# ─────────────────────────────────────────────────────────────────────────────
# Gerador de Documentos Veterinários — Isabelle Rizzo Assumpção
# ─────────────────────────────────────────────────────────────────────────────
# Templates: Dieta Caseira | Ração | Prescrição | Solicitação de Exames
# ─────────────────────────────────────────────────────────────────────────────

import streamlit as st
import io, os, re, tempfile, subprocess, platform, base64, json
from datetime import datetime
from pathlib import Path

from docx import Document
from PIL import Image
import pytesseract
from pypdf import PdfWriter, PdfReader
import anthropic

from fill_template import (
    fill_template   as _fill_dieta,
    fill_exames     as _fill_exames,
    fill_prescricao as _fill_prescricao,
)

# ── Tesseract (Windows) ───────────────────────────────────────────────────────
if platform.system() == "Windows":
    for _tp in [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    ]:
        if os.path.exists(_tp):
            pytesseract.pytesseract.tesseract_cmd = _tp
            break

# ── Constantes ────────────────────────────────────────────────────────────────
MESES = [
    "JANEIRO","FEVEREIRO","MARÇO","ABRIL","MAIO","JUNHO",
    "JULHO","AGOSTO","SETEMBRO","OUTUBRO","NOVEMBRO","DEZEMBRO",
]

TEMPLATES = {
    "🥩 Dieta Caseira":         {"path": "templates/template_dieta_caseira.docx", "tipo": "dieta"},
    "🐾 Ração":                 {"path": "templates/template_dieta_racao.docx",   "tipo": "dieta"},
    "💊 Prescrição":            {"path": "templates/template_prescricao.docx",    "tipo": "prescricao"},
    "🔬 Solicitação de Exames": {"path": "templates/template_exames.docx",        "tipo": "exames"},
}

VIAS = ["USO ORAL", "USO TÓPICO", "USO OTOLÓGICO", "USO SUBCUTÂNEO"]


# ═══════════════════════════════════════════════════════════════════════════════
# 1. PARSING DO ANAMNESE
# ═══════════════════════════════════════════════════════════════════════════════

def parse_anamnese(docx_bytes: bytes) -> dict:
    doc = Document(io.BytesIO(docx_bytes))
    paras = [p.text.strip() for p in doc.paragraphs if p.text.strip()]

    data = dict(
        tutor_nome="", tutor_cpf="", tutor_email="", tutor_endereco="",
        pet_nome="", pet_especie="", pet_raca="", pet_sexo="", pet_nascimento="",
    )

    pet_idx = next(
        (i for i, p in enumerate(paras) if re.match(r"^pet\s*:", p, re.I)), -1
    )

    for p in (paras[:pet_idx] if pet_idx >= 0 else paras):
        clean = p.strip()
        digits = re.sub(r"[\.\-\/\s]", "", clean)
        if re.match(r"^\d{11}$", digits):
            data["tutor_cpf"] = clean
        elif "@" in clean:
            data["tutor_email"] = clean
        elif re.search(r"\b(rua|av\.|avenida|alameda|travessa|estrada|rod\.)\b", clean, re.I):
            data["tutor_endereco"] = clean
        elif not data["tutor_nome"] and len(clean.split()) >= 2:
            data["tutor_nome"] = clean

    if pet_idx >= 0:
        pets = paras[pet_idx + 1:]
        positional = [p for p in pets if not re.match(r"^(microchip|cor\s|peso\s|nasc:)", p, re.I)]
        if len(positional) > 0: data["pet_nome"]      = positional[0].strip()
        if len(positional) > 1: data["pet_raca"]      = positional[1].strip()
        if len(positional) > 2: data["pet_sexo"]      = positional[2].strip()
        for p in pets:
            m = re.match(r"nasc[:\s]+(.+)", p, re.I)
            if m:
                data["pet_nascimento"] = m.group(1).strip()

    return data


# ═══════════════════════════════════════════════════════════════════════════════
# 2. OCR DO DIETALABS (Claude Vision)
# ═══════════════════════════════════════════════════════════════════════════════

def ocr_dietalabs(image_bytes: bytes, api_key: str = "") -> tuple[list[tuple[str, str]], str]:
    if not api_key:
        raise ValueError("Chave de API Anthropic não configurada.")

    img = Image.open(io.BytesIO(image_bytes))
    fmt = (img.format or "JPEG").upper()
    media_map = {"JPEG": "image/jpeg", "JPG": "image/jpeg", "PNG": "image/png",
                 "WEBP": "image/webp", "GIF": "image/gif"}
    media_type = media_map.get(fmt, "image/jpeg")
    img_b64 = base64.standard_b64encode(image_bytes).decode("utf-8")

    prompt = (
        "Você está vendo um print do sistema Dietalabs de nutrição veterinária.\n\n"
        "Extraia a lista de ingredientes com suas quantidades em gramas e o total.\n\n"
        "Retorne APENAS um JSON válido neste formato exato (sem texto antes ou depois, sem markdown):\n"
        '{"ingredientes": [{"nome": "Nome do ingrediente", "quantidade": "100"}, ...], "total": "337"}\n\n'
        "Regras:\n"
        "- quantidade: apenas o número, sem 'g' ou unidade\n"
        "- total: apenas o número, sem 'g'\n"
        "- Ignore o cabeçalho 'Lista de ingredientes'\n"
        "- Ignore o botão × ao lado de cada ingrediente"
    )

    client = anthropic.Anthropic(api_key=api_key)
    message = client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=1024,
        messages=[{"role": "user", "content": [
            {"type": "image", "source": {"type": "base64", "media_type": media_type, "data": img_b64}},
            {"type": "text", "text": prompt},
        ]}],
    )

    raw = message.content[0].text.strip()
    raw = re.sub(r"```(?:json)?\n?", "", raw).strip().rstrip("`").strip()
    data = json.loads(raw)

    ingredients = [
        (item["nome"].strip(), str(item["quantidade"]).strip())
        for item in data.get("ingredientes", [])
        if item.get("nome") and item.get("quantidade")
    ]
    total = str(data.get("total", "")).strip()
    return ingredients, total


# ═══════════════════════════════════════════════════════════════════════════════
# 3. GERAÇÃO DE PDF
# ═══════════════════════════════════════════════════════════════════════════════

def _data_extenso() -> str:
    hoje = datetime.now()
    return f"{hoje.day} DE {MESES[hoje.month - 1]} DE {hoje.year}"


def _soffice_path() -> str:
    if platform.system() == "Windows":
        for c in [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]:
            if os.path.exists(c): return c
    return "libreoffice"


def convert_to_pdf(docx_bytes: bytes) -> bytes:
    soffice = _soffice_path()
    with tempfile.TemporaryDirectory() as tmp:
        docx_path = os.path.join(tmp, "input.docx")
        pdf_path  = os.path.join(tmp, "input.pdf")
        with open(docx_path, "wb") as f:
            f.write(docx_bytes)
        result = subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf", "--outdir", tmp, docx_path],
            capture_output=True, timeout=120,
        )
        if result.returncode != 0:
            raise RuntimeError(f"LibreOffice erro:\n{result.stderr.decode(errors='replace')}")
        if not os.path.exists(pdf_path):
            raise RuntimeError("LibreOffice não gerou o PDF esperado.")
        with open(pdf_path, "rb") as f:
            return f.read()


def generate_dieta_pdf(template_path: str, data: dict, opcoes: list) -> bytes:
    data_com_data = dict(data)
    data_com_data["data"] = _data_extenso()

    pdfs_opcoes = []
    for i, (ingredients, _total) in enumerate(opcoes):
        with st.spinner(f"Gerando Opção {i + 1} de {len(opcoes)}..."):
            docx_bytes = _fill_dieta(template_path, data_com_data, ingredients, str(i + 1))
            pdfs_opcoes.append(convert_to_pdf(docx_bytes))

    writer = PdfWriter()
    for pdf_bytes in pdfs_opcoes:
        reader = PdfReader(io.BytesIO(pdf_bytes))
        writer.append(reader, pages=(0, 1))
    if pdfs_opcoes:
        reader = PdfReader(io.BytesIO(pdfs_opcoes[0]))
        n = len(reader.pages)
        if n > 1:
            writer.append(reader, pages=(1, n))

    output = io.BytesIO()
    writer.write(output)
    return output.getvalue()


def generate_exames_pdf(template_path: str, data: dict, exames: list) -> bytes:
    data_com_data = dict(data)
    data_com_data["data"] = _data_extenso()
    with st.spinner("Gerando solicitação de exames..."):
        return convert_to_pdf(_fill_exames(template_path, data_com_data, exames))


def generate_prescricao_pdf(template_path: str, data: dict, vias_data: list) -> bytes:
    data_com_data = dict(data)
    data_com_data["data"] = _data_extenso()
    with st.spinner("Gerando prescrição..."):
        return convert_to_pdf(_fill_prescricao(template_path, data_com_data, [vias_data]))


# ═══════════════════════════════════════════════════════════════════════════════
# 4. UI — PRESCRIÇÃO
# ═══════════════════════════════════════════════════════════════════════════════

def _via_key(via: str) -> str:
    return via.replace(" ", "_")


def _init_prescricao():
    for via in VIAS:
        k = f"presc_count_{_via_key(via)}"
        if k not in st.session_state:
            st.session_state[k] = 1


def _ui_prescricao() -> list[dict]:
    st.caption("Preencha as seções que se aplicam. Seções vazias serão ignoradas.")

    for via in VIAS:
        vk = _via_key(via)
        count_key = f"presc_count_{vk}"
        count = st.session_state[count_key]

        # Verificar se algum campo já foi preenchido (para expandir automaticamente)
        has_content = any(
            st.session_state.get(f"presc_{vk}_{i}_item", "").strip()
            for i in range(count)
        )

        with st.expander(f"💊 {via}", expanded=(via == "USO ORAL" or has_content)):
            for i in range(count):
                item_key  = f"presc_{vk}_{i}_item"
                instr_key = f"presc_{vk}_{i}_instr"

                col1, col2, col_del = st.columns([3, 4, 0.5])
                with col1:
                    st.text_input(
                        f"Medicamento {i + 1}",
                        key=item_key,
                        placeholder="Ex: Mirtazapina 1,88mg — Fórmula magistral",
                    )
                with col2:
                    st.text_area(
                        f"Instrução {i + 1}",
                        key=instr_key,
                        height=90,
                        placeholder="Ex: Administrar 1 comprimido por via oral a cada 72h por 30 dias.",
                    )
                with col_del:
                    st.markdown("<br><br>", unsafe_allow_html=True)
                    if count > 1 and st.button("🗑️", key=f"del_{vk}_{i}", help="Remover"):
                        # Desloca os valores para cima
                        for j in range(i, count - 1):
                            st.session_state[f"presc_{vk}_{j}_item"]  = st.session_state.get(f"presc_{vk}_{j+1}_item", "")
                            st.session_state[f"presc_{vk}_{j}_instr"] = st.session_state.get(f"presc_{vk}_{j+1}_instr", "")
                        st.session_state.pop(f"presc_{vk}_{count-1}_item",  None)
                        st.session_state.pop(f"presc_{vk}_{count-1}_instr", None)
                        st.session_state[count_key] -= 1
                        st.rerun()

            if st.button(f"+ Adicionar medicamento", key=f"add_{vk}"):
                st.session_state[count_key] += 1
                st.rerun()

    # Coletar resultado final
    vias_data = []
    for via in VIAS:
        vk    = _via_key(via)
        count = st.session_state[f"presc_count_{vk}"]
        meds  = []
        for i in range(count):
            item  = st.session_state.get(f"presc_{vk}_{i}_item",  "").strip()
            instr = st.session_state.get(f"presc_{vk}_{i}_instr", "").strip()
            if item:
                meds.append({"item": item, "instrucao": instr})
        if meds:
            vias_data.append({"via": via, "medicamentos": meds})

    return vias_data


# ═══════════════════════════════════════════════════════════════════════════════
# 5. MAIN
# ═══════════════════════════════════════════════════════════════════════════════

def main():
    st.set_page_config(
        page_title="Documentos • Isabelle Rizzo",
        page_icon="🐾",
        layout="centered",
    )

    _init_prescricao()

    st.title("🐾 Gerador de Documentos")
    st.caption("Isabelle Rizzo Assumpção • CRMV 48652/SP")

    # ── Sidebar ───────────────────────────────────────────────────────────────
    with st.sidebar:
        st.header("⚙️ Configurações")
        template_key  = st.selectbox("Tipo de documento", list(TEMPLATES.keys()))
        template_info = TEMPLATES[template_key]
        template_path = template_info["path"]
        tipo          = template_info["tipo"]
        st.markdown("---")

        _env_key = os.environ.get("ANTHROPIC_API_KEY", "")
        if _env_key:
            api_key = _env_key
            st.success("🔑 Chave de API carregada do ambiente.")
        else:
            api_key = st.text_input(
                "🔑 Chave API Anthropic",
                type="password",
                help="Necessária para leitura automática dos prints do Dietalabs.",
            )
            if not api_key:
                st.warning("Sem chave de API: preencha os ingredientes manualmente.")

        st.markdown("---")
        st.caption("Apsirtus Clínica Veterinária")

    # ── Passo 1: Ficha do cliente ─────────────────────────────────────────────
    st.subheader("1. Ficha do cliente")
    docx_file = st.file_uploader("Envie o arquivo ANAMNESE.docx", type=["docx"])

    anamnese: dict = {}
    if docx_file:
        try:
            anamnese = parse_anamnese(docx_file.read())
            st.success("✅ Ficha lida. Confira os dados abaixo e corrija se necessário:")
        except Exception as e:
            st.error(f"Erro ao ler o arquivo Word: {e}")

        col1, col2 = st.columns(2)
        with col1:
            st.markdown("**Animal**")
            anamnese["pet_nome"]       = st.text_input("Nome do pet",       anamnese.get("pet_nome", ""))
            anamnese["pet_especie"]    = st.text_input("Espécie",            anamnese.get("pet_especie", "Cão"))
            anamnese["pet_raca"]       = st.text_input("Raça",               anamnese.get("pet_raca", ""))
            anamnese["pet_sexo"]       = st.text_input("Sexo",               anamnese.get("pet_sexo", ""))
            anamnese["pet_nascimento"] = st.text_input("Data de nascimento", anamnese.get("pet_nascimento", ""))
        with col2:
            st.markdown("**Responsável**")
            anamnese["tutor_nome"]     = st.text_input("Nome do tutor", anamnese.get("tutor_nome", ""))
            anamnese["tutor_cpf"]      = st.text_input("CPF",           anamnese.get("tutor_cpf", ""))
            anamnese["tutor_endereco"] = st.text_input("Endereço",      anamnese.get("tutor_endereco", ""))

    # ── Passo 2: Conteúdo (varia por tipo) ────────────────────────────────────
    st.subheader("2. Conteúdo")

    opcoes:    list = []
    exames:    list = []
    vias_data: list = []
    conteudo_ok = False

    # ── DIETA ─────────────────────────────────────────────────────────────────
    if tipo == "dieta":
        st.caption("Envie um ou mais prints do Dietalabs — cada um será uma Opção no plano.")
        imagens = st.file_uploader(
            "Prints (PNG ou JPG)",
            type=["png", "jpg", "jpeg"],
            accept_multiple_files=True,
        )

        if imagens:
            for i, img_file in enumerate(imagens):
                with st.expander(f"Opção {i + 1} — {img_file.name}", expanded=True):
                    col_img, col_data = st.columns([1, 2])
                    img_bytes = img_file.read()

                    with col_img:
                        st.image(img_bytes, use_container_width=True)

                    try:
                        ingredients, total = ocr_dietalabs(img_bytes, api_key=api_key)
                        ocr_ok = True
                    except Exception as e:
                        ingredients, total = [], ""
                        ocr_ok = False
                        st.warning(f"Leitura automática indisponível ({e}). Preencha manualmente.")

                    raw = "\n".join(f"{n} | {q}" for n, q in ingredients)
                    if total:
                        raw += f"\nTOTAL | {total}"

                    with col_data:
                        st.markdown("**Ingredientes** — formato: `Nome | Quantidade (g)`")
                        if ocr_ok and ingredients:
                            st.caption("✅ Extraído automaticamente. Corrija se necessário.")
                        else:
                            st.caption("✏️ Digite os ingredientes manualmente.")

                        edited = st.text_area(
                            "Ingredientes",
                            value=raw,
                            height=220,
                            key=f"ing_{i}",
                            label_visibility="collapsed",
                        )

                    parsed_ings = []
                    parsed_total = total
                    for line in edited.splitlines():
                        if "|" not in line:
                            continue
                        parts  = line.split("|", 1)
                        nome_e = parts[0].strip()
                        qtd_e  = parts[1].strip()
                        if nome_e.upper() == "TOTAL":
                            parsed_total = qtd_e
                        elif nome_e:
                            parsed_ings.append((nome_e, qtd_e))

                    opcoes.append((parsed_ings, parsed_total))

            conteudo_ok = bool(opcoes)

    # ── EXAMES ────────────────────────────────────────────────────────────────
    elif tipo == "exames":
        st.caption("Digite um exame por linha.")
        exames_raw = st.text_area(
            "Exames solicitados",
            height=220,
            placeholder="Hemograma completo\nBioquímico renal (ureia, creatinina, ALT, FA)\nUrinálise (EAS)\n...",
        )
        exames = [line.strip() for line in exames_raw.splitlines() if line.strip()]
        conteudo_ok = bool(exames)
        if not exames:
            st.info("Digite pelo menos um exame para habilitar a geração.")

    # ── PRESCRIÇÃO ────────────────────────────────────────────────────────────
    elif tipo == "prescricao":
        vias_data   = _ui_prescricao()
        conteudo_ok = bool(vias_data)
        if not vias_data:
            st.info("Preencha pelo menos um medicamento em uma das seções acima.")

    # ── Passo 3: Gerar PDF ────────────────────────────────────────────────────
    st.subheader("3. Gerar documento")

    ready = bool(docx_file and anamnese and conteudo_ok)
    if not ready:
        st.info("Preencha a ficha do cliente e o conteúdo acima para habilitar a geração.")

    if st.button("📄 Gerar PDF", type="primary", disabled=not ready):
        if not os.path.exists(template_path):
            st.error(f"Template não encontrado: `{template_path}`")
        else:
            try:
                if tipo == "dieta":
                    pdf_bytes = generate_dieta_pdf(template_path, anamnese, opcoes)
                elif tipo == "exames":
                    pdf_bytes = generate_exames_pdf(template_path, anamnese, exames)
                elif tipo == "prescricao":
                    pdf_bytes = generate_prescricao_pdf(template_path, anamnese, vias_data)

                nome_pet  = anamnese.get("pet_nome", "pet").replace(" ", "_")
                tipo_nome = template_key.split()[-1]
                filename  = f"{tipo_nome}_{nome_pet}_{datetime.now().strftime('%d%m%Y')}.pdf"

                st.success("✅ PDF gerado com sucesso!")
                st.download_button(
                    label="⬇️ Baixar PDF",
                    data=pdf_bytes,
                    file_name=filename,
                    mime="application/pdf",
                )
            except Exception as e:
                st.error(f"Erro ao gerar o PDF: {e}")
                with st.expander("Detalhes do erro"):
                    st.exception(e)


if __name__ == "__main__":
    main()
