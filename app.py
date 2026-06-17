# ─────────────────────────────────────────────────────────────────────────────
# Gerador de Documentos Veterinários — Isabelle Rizzo Assumpção
# ─────────────────────────────────────────────────────────────────────────────
# Templates: Dieta Caseira | Ração | Prescrição | Solicitação de Exames
# ─────────────────────────────────────────────────────────────────────────────

import streamlit as st
import io, os, re, tempfile, subprocess, platform, base64, json, gc
from datetime import datetime
from pathlib import Path

from docx import Document
from PIL import Image
import pytesseract
from pypdf import PdfWriter, PdfReader, Transformation
import anthropic

from fill_template import (
    fill_template                        as _fill_dieta,
    fill_exames                          as _fill_exames,
    fill_prescricao                      as _fill_prescricao,
    build_eliminacao_instructions_docx   as _build_eliminacao,
)
from draft import save_draft, draft_info, clear_draft

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
    "🥩 Dieta Caseira":         {"path": "templates/template_dieta_caseira.docx", "tipo": "dieta", "subtipo": "caseira"},
    "🐾 Ração":                 {"path": "templates/template_dieta_racao.docx",   "tipo": "dieta", "subtipo": "racao"},
    "💊 Prescrição":            {"path": "templates/template_prescricao.docx",    "tipo": "prescricao"},
    "🔬 Solicitação de Exames": {"path": "templates/template_exames.docx",        "tipo": "exames"},
}

VIAS = ["USO ORAL", "USO TÓPICO", "USO OTOLÓGICO", "USO SUBCUTÂNEO"]

ANAM_FIELDS = [
    "pet_nome", "pet_especie", "pet_raca", "pet_sexo", "pet_nascimento",
    "tutor_nome", "tutor_cpf", "tutor_endereco",
]


# ── Monitoramento de boot e memória (diagnóstico de reset) ───────────────────
def _rss_mb() -> int:
    """Uso de memória residente do processo, em MB (-1 se indisponível)."""
    try:
        with open("/proc/self/status") as f:
            for line in f:
                if line.startswith("VmRSS:"):
                    return int(line.split()[1]) // 1024
    except Exception:
        pass
    try:
        import resource
        return resource.getrusage(resource.RUSAGE_SELF).ru_maxrss // 1024
    except Exception:
        return -1


@st.cache_resource
def _log_boot():
    """Marcador único por processo — vira a linha de 'boot' no log do Manage app."""
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print("=" * 64, flush=True)
    print(f"APP INICIOU (boot) — {ts} — RAM inicial: {_rss_mb()} MB", flush=True)
    print("=" * 64, flush=True)
    return True


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

    def val(text: str) -> str:
        """Extrai o valor após ':' numa linha 'Label: valor'."""
        if ":" in text:
            return text.split(":", 1)[1].strip()
        return text.strip()

    in_pet = False
    for p in paras:
        pl = p.lower()

        if "nome do pet" in pl:
            in_pet = True
            data["pet_nome"] = val(p)
        elif not in_pet:
            if "nome completo" in pl or ("nome" in pl and "tutor" in pl):
                data["tutor_nome"] = val(p)
            elif pl.startswith("cpf"):
                data["tutor_cpf"] = val(p)
            elif "e-mail" in pl or "email" in pl:
                data["tutor_email"] = val(p)
            elif "endereço" in pl or "endereco" in pl:
                data["tutor_endereco"] = val(p)
        else:  # seção do pet
            if "espécie" in pl or "especie" in pl:
                data["pet_especie"] = val(p)
            elif "raça" in pl or "raca" in pl:
                data["pet_raca"] = val(p)
            elif pl.startswith("sexo"):
                data["pet_sexo"] = val(p)
            elif "data de nascimento" in pl:
                data["pet_nascimento"] = val(p)

    return data


# ═══════════════════════════════════════════════════════════════════════════════
# 2. OCR DO DIETALABS (Claude Vision)
# ═══════════════════════════════════════════════════════════════════════════════

def ocr_dietalabs(image_bytes: bytes, api_key: str = "") -> tuple[list[tuple[str, str]], str]:
    if not api_key:
        raise ValueError("Chave de API Anthropic não configurada.")

    # Reduz a imagem antes de enviar: corta o uso de memória e o tamanho do
    # payload da API, mantendo resolução de sobra para o OCR de texto.
    img = Image.open(io.BytesIO(image_bytes)).convert("RGB")
    MAX_LADO = 1568   # mesmo limite que a API de visão da Anthropic já aplica internamente
    if max(img.size) > MAX_LADO:
        img.thumbnail((MAX_LADO, MAX_LADO), Image.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=90)
    img.close()
    media_type = "image/jpeg"
    img_b64 = base64.standard_b64encode(buf.getvalue()).decode("utf-8")
    buf.close()
    del img, buf
    gc.collect()

    prompt = (
        "Você está vendo um print do sistema Dietalabs de nutrição veterinária.\n\n"
        "Extraia a lista de ingredientes com suas quantidades e o total.\n\n"
        "Retorne APENAS um JSON válido neste formato exato (sem texto antes ou depois, sem markdown):\n"
        '{"ingredientes": [{"nome": "Nome do ingrediente", "quantidade": "50 gramas"}, ...], "total": "337 gramas"}\n\n'
        "Regras:\n"
        "- quantidade: número seguido da unidade escrita por extenso\n"
        "  Exemplos: '50 gramas', '100 gramas', '4 ml', '1 grama'\n"
        "  Use 'grama' (singular) para valor 1, 'gramas' para os demais\n"
        "  Se a unidade for ml, mantenha 'ml'\n"
        "- total: número seguido de 'gramas' (ex: '337 gramas')\n"
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


def _stamp_signature_on_pdf(pdf_bytes: bytes) -> bytes:
    """Adiciona assinatura digital diretamente no PDF via PIL+pypdf.

    Funciona independente do LibreOffice — sobrepõe a imagem PNG como stamp
    em cada página do PDF na posição medida do rodapé.
    """
    sig_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "assinatura.png")
    if not os.path.exists(sig_path):
        return pdf_bytes
    try:
        # Dimensões do stamp (medidas do rodapé gerado localmente)
        SIG_W_PT = 108          # 1.5 inches em pontos PDF
        SIG_H_PT = int(SIG_W_PT * 200 / 710)  # ~30 pt (proporcional 710×200)
        X_PT     = 57           # margem esquerda (~2 cm)
        Y_PT     = 88           # base da assinatura a partir do fundo da página

        # Redimensionar PNG para o tamanho alvo (72 DPI → 1pt = 1px)
        sig_img = Image.open(sig_path).convert("RGB")
        sig_resized = sig_img.resize((SIG_W_PT, SIG_H_PT), Image.LANCZOS)

        # Salvar assinatura como mini-PDF (página = tamanho da imagem)
        sig_pdf_buf = io.BytesIO()
        sig_resized.save(sig_pdf_buf, format="PDF", resolution=72)
        sig_pdf_buf.seek(0)

        # Carregar, posicionar e expandir para A4
        stamp_reader = PdfReader(sig_pdf_buf)
        stamp_page   = stamp_reader.pages[0]
        stamp_page.add_transformation(Transformation().translate(X_PT, Y_PT))
        stamp_page.mediabox.lower_left  = (0, 0)
        stamp_page.mediabox.upper_right = (595.28, 841.89)

        # Mesclar stamp em cada página
        reader = PdfReader(io.BytesIO(pdf_bytes))
        writer = PdfWriter()
        for page in reader.pages:
            page.merge_page(stamp_page)
            writer.add_page(page)

        out = io.BytesIO()
        writer.write(out)
        return out.getvalue()
    except Exception:
        return pdf_bytes  # fallback: retorna PDF sem assinatura


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
            pdf = f.read()
    return _stamp_signature_on_pdf(pdf)


def generate_dieta_pdf(template_path: str, data: dict, opcoes: list,
                       eliminacao: bool = False, subtipo: str = "caseira") -> bytes:
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

    if eliminacao:
        with st.spinner("Gerando instruções de dieta de eliminação..."):
            instr_pdf = convert_to_pdf(_build_eliminacao(subtipo, data_com_data.get("data", "")))
        writer.append(PdfReader(io.BytesIO(instr_pdf)))
    elif pdfs_opcoes:
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

    _log_boot()
    _init_prescricao()

    st.title("🐾 Gerador de Documentos")
    st.caption("Isabelle Rizzo Assumpção • CRMV 48652/SP")

    # ── Recuperação de rascunho ───────────────────────────────────────────
    _info = draft_info()
    if _info and not st.session_state.get("_rascunho_resolvido"):
        st.warning(
            f"💾 Há um rascunho salvo em **{_info['when']}**. "
            "Quer recuperar o que estava preenchido?"
        )
        _c1, _c2, _ = st.columns([1, 1, 3])
        if _c1.button("♻️ Recuperar"):
            for _k, _v in _info["fields"].items():
                st.session_state[_k] = _v
            st.session_state["_rascunho_resolvido"] = True
            st.rerun()
        if _c2.button("🗑️ Descartar"):
            clear_draft()
            for _k in list(st.session_state.keys()):
                if _k.startswith(("anam_", "presc_", "opt_")) or _k == "exames_raw":
                    del st.session_state[_k]
            st.session_state["_rascunho_resolvido"] = True
            st.rerun()

    # ── Sidebar ───────────────────────────────────────────────────────────────
    with st.sidebar:
        st.header("⚙️ Configurações")
        template_key  = st.selectbox("Tipo de documento", list(TEMPLATES.keys()))
        template_info = TEMPLATES[template_key]
        template_path = template_info["path"]
        tipo          = template_info["tipo"]
        subtipo       = template_info.get("subtipo", "")
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
        st.caption("Apsirtos Clínica Veterinária")

    # ── Passo 1: Ficha do cliente ─────────────────────────────────────────────
    st.subheader("1. Ficha do cliente")
    docx_file = st.file_uploader("Envie o arquivo ANAMNESE.docx", type=["docx"])

    if docx_file:
        try:
            parsed = parse_anamnese(docx_file.read())
            # Semeia os campos só uma vez por arquivo (não apaga edições/rascunho)
            if st.session_state.get("_docx_semeado") != docx_file.name:
                for _k in ANAM_FIELDS:
                    st.session_state[f"anam_{_k}"] = parsed.get(_k, "")
                if not st.session_state.get("anam_pet_especie"):
                    st.session_state["anam_pet_especie"] = "Cão"
                st.session_state["_docx_semeado"] = docx_file.name
            st.success("✅ Ficha lida. Confira os dados abaixo e corrija se necessário:")
        except Exception as e:
            st.error(f"Erro ao ler o arquivo Word: {e}")

    tem_ficha = bool(docx_file) or any(
        st.session_state.get(f"anam_{_k}", "") for _k in ANAM_FIELDS
    )

    anamnese: dict = {}
    if tem_ficha:
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("**Animal**")
            st.text_input("Nome do pet",       key="anam_pet_nome")
            st.text_input("Espécie",            key="anam_pet_especie")
            st.text_input("Raça",               key="anam_pet_raca")
            st.text_input("Sexo",               key="anam_pet_sexo")
            st.text_input("Data de nascimento", key="anam_pet_nascimento")
        with col2:
            st.markdown("**Responsável**")
            st.text_input("Nome do tutor", key="anam_tutor_nome")
            st.text_input("CPF",           key="anam_tutor_cpf")
            st.text_input("Endereço",      key="anam_tutor_endereco")
        anamnese = {_k: st.session_state.get(f"anam_{_k}", "") for _k in ANAM_FIELDS}

    # ── Passo 2: Conteúdo (varia por tipo) ────────────────────────────────────
    st.subheader("2. Conteúdo")

    opcoes:    list = []
    exames:    list = []
    vias_data: list = []
    eliminacao: bool  = False
    conteudo_ok = False

    # ── DIETA ─────────────────────────────────────────────────────────────────
    if tipo == "dieta":
        eliminacao = st.checkbox(
            "🔬 Dieta de Eliminação",
            key="opt_eliminacao",
            help="Substitui as instruções do documento pelas orientações específicas de dieta de eliminação.",
        )
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
                        st.image(img_bytes, width="stretch")

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
            key="exames_raw",
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

    ready = bool(anamnese.get("pet_nome") and conteudo_ok)
    if not ready:
        st.info("Preencha a ficha do cliente e o conteúdo acima para habilitar a geração.")

    if st.button("📄 Gerar PDF", type="primary", disabled=not ready):
        if not os.path.exists(template_path):
            st.error(f"Template não encontrado: `{template_path}`")
        else:
            try:
                if tipo == "dieta":
                    pdf_bytes = generate_dieta_pdf(template_path, anamnese, opcoes,
                                                       eliminacao=eliminacao, subtipo=subtipo)
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

    # ── Autosave do rascunho (sempre por último) ──────────────────────────
    _snap = {
        k: st.session_state[k]
        for k in list(st.session_state.keys())
        if k.startswith(("anam_", "presc_", "opt_")) or k == "exames_raw"
    }
    _tem_conteudo = (
        bool(str(_snap.get("anam_pet_nome", "")).strip())
        or bool(str(_snap.get("exames_raw", "")).strip())
        or any(k.endswith("_item") and str(v).strip() for k, v in _snap.items())
    )
    if _tem_conteudo:
        save_draft(_snap)

    # ── Monitor de memória: só alerta perto do limite (~1GB do plano grátis) ──
    _rss = _rss_mb()
    if _rss >= 700:
        print(
            f"RAM ALTA: {_rss} MB às {datetime.now().strftime('%H:%M:%S')} "
            "(limite ~1024 MB no plano grátis)",
            flush=True,
        )


if __name__ == "__main__":
    main()
