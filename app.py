# ─────────────────────────────────────────────────────────────────────────────
# Gerador de Plano Alimentar Veterinário — Isabelle Rizzo Assumpção
# ─────────────────────────────────────────────────────────────────────────────
# Inputs:  ANAMNESE.docx  +  prints do Dietalabs  +  escolha de template
# Output:  PDF A4 pronto para envio ao cliente
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

from fill_template import fill_template as _fill_docx

# ── Configurar caminho do Tesseract no Windows ────────────────────────────────
if platform.system() == "Windows":
    for _tp in [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    ]:
        if os.path.exists(_tp):
            pytesseract.pytesseract.tesseract_cmd = _tp
            break

# ── Constantes ─────────────────────────────────────────────────────────────
MESES = [
    "JANEIRO","FEVEREIRO","MARÇO","ABRIL","MAIO","JUNHO",
    "JULHO","AGOSTO","SETEMBRO","OUTUBRO","NOVEMBRO","DEZEMBRO",
]

TEMPLATES = {
    "🥩 Dieta Caseira": "templates/template_dieta_caseira.docx",
    "🐾 Ração":         "templates/template_dieta_racao.docx",
}


# ═══════════════════════════════════════════════════════════════════════════════
# 1. PARSING DO ANAMNESE
# ═══════════════════════════════════════════════════════════════════════════════

def parse_anamnese(docx_bytes: bytes) -> dict:
    """Extrai campos do ANAMNESE.docx e retorna dicionário estruturado."""
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
        positional = [
            p for p in pets
            if not re.match(r"^(microchip|cor\s|peso\s|nasc:)", p, re.I)
        ]
        if len(positional) > 0: data["pet_nome"]  = positional[0].strip()
        if len(positional) > 1: data["pet_raca"]  = positional[1].strip()
        if len(positional) > 2: data["pet_sexo"]  = positional[2].strip()

        for p in pets:
            m = re.match(r"nasc[:\s]+(.+)", p, re.I)
            if m:
                data["pet_nascimento"] = m.group(1).strip()

    return data


# ═══════════════════════════════════════════════════════════════════════════════
# 2. OCR DO DIETALABS
# ═══════════════════════════════════════════════════════════════════════════════

def ocr_dietalabs(image_bytes: bytes, api_key: str = "") -> tuple[list[tuple[str, str]], str]:
    """
    Lê um print do Dietalabs usando Claude Vision e retorna:
      - lista de tuplas (nome_ingrediente, quantidade_str)  ex.: ("Frango cozido", "150")
      - total_str  ex.: "337"
    """
    if not api_key:
        raise ValueError("Chave de API Anthropic não configurada.")

    # Detectar tipo da imagem
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
        "- Ignore a linha de cabeçalho 'Lista de ingredientes'\n"
        "- Ignore o botão × ao lado de cada ingrediente"
    )

    client = anthropic.Anthropic(api_key=api_key)
    message = client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=1024,
        messages=[{
            "role": "user",
            "content": [
                {
                    "type": "image",
                    "source": {
                        "type": "base64",
                        "media_type": media_type,
                        "data": img_b64,
                    },
                },
                {"type": "text", "text": prompt},
            ],
        }],
    )

    raw = message.content[0].text.strip()

    # Remover blocos markdown se o modelo os incluir
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
# 3. GERAÇÃO DO PDF
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
            if os.path.exists(c):
                return c
    return "libreoffice"


def convert_to_pdf(docx_bytes: bytes) -> bytes:
    """Converte DOCX (bytes) → PDF (bytes) via LibreOffice headless."""
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
            raise RuntimeError(
                f"LibreOffice retornou erro:\n{result.stderr.decode(errors='replace')}"
            )
        if not os.path.exists(pdf_path):
            raise RuntimeError("LibreOffice não gerou o arquivo PDF esperado.")

        with open(pdf_path, "rb") as f:
            return f.read()


def generate_final_pdf(template_path: str, data: dict, opcoes: list[tuple]) -> bytes:
    """
    Gera PDF final com N opções de dieta + páginas de instrução.

    Estrutura:
      • Página 1 .. N  → uma página de dieta por opção
      • Páginas N+1 .. → páginas de instrução (do template)
    """
    data_com_data = dict(data)
    data_com_data["data"] = _data_extenso()

    pdfs_opcoes: list[bytes] = []
    for i, (ingredients, total) in enumerate(opcoes):
        with st.spinner(f"Gerando Opção {i + 1} de {len(opcoes)}..."):
            docx_bytes = _fill_docx(template_path, data_com_data, ingredients, total, str(i + 1))
            pdf_bytes  = convert_to_pdf(docx_bytes)
            pdfs_opcoes.append(pdf_bytes)

    writer = PdfWriter()

    # Página de dieta de cada opção (somente a 1ª página de cada PDF)
    for pdf_bytes in pdfs_opcoes:
        reader = PdfReader(io.BytesIO(pdf_bytes))
        writer.append(reader, pages=(0, 1))

    # Páginas de instrução do 1º PDF (páginas 2 em diante)
    if pdfs_opcoes:
        reader = PdfReader(io.BytesIO(pdfs_opcoes[0]))
        n = len(reader.pages)
        if n > 1:
            writer.append(reader, pages=(1, n))

    output = io.BytesIO()
    writer.write(output)
    return output.getvalue()


# ═══════════════════════════════════════════════════════════════════════════════
# 4. INTERFACE STREAMLIT
# ═══════════════════════════════════════════════════════════════════════════════

def main():
    st.set_page_config(
        page_title="Plano Alimentar • Isabelle Rizzo",
        page_icon="🐾",
        layout="centered",
    )

    st.title("🐾 Gerador de Plano Alimentar")
    st.caption("Isabelle Rizzo Assumpção • CRMV 48652/SP")

    # ── Sidebar: escolha do template + chave de API ───────────────────────────
    with st.sidebar:
        st.header("⚙️ Configurações")
        template_key  = st.selectbox("Tipo de dieta", list(TEMPLATES.keys()))
        template_path = TEMPLATES[template_key]
        st.markdown("---")

        # Chave de API: variável de ambiente tem prioridade
        _env_key = os.environ.get("ANTHROPIC_API_KEY", "")
        if _env_key:
            api_key = _env_key
            st.success("🔑 Chave de API carregada do ambiente.")
        else:
            api_key = st.text_input(
                "🔑 Chave API Anthropic",
                type="password",
                help="Necessária para leitura automática dos prints do Dietalabs. "
                     "Obtenha em console.anthropic.com",
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

    # ── Passo 2: Prints do Dietalabs ─────────────────────────────────────────
    st.subheader("2. Prints do Dietalabs")
    st.caption("Envie um ou mais prints — cada um será uma Opção no plano.")

    imagens = st.file_uploader(
        "Prints (PNG ou JPG)",
        type=["png", "jpg", "jpeg"],
        accept_multiple_files=True,
    )

    opcoes: list[tuple] = []

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

                parsed_ings: list[tuple[str, str]] = []
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

    # ── Passo 3: Gerar PDF ────────────────────────────────────────────────────
    st.subheader("3. Gerar PDF")

    ready = bool(docx_file and anamnese and imagens and opcoes)
    if not ready:
        st.info("Preencha a ficha do cliente e envie pelo menos um print do Dietalabs.")

    if st.button("📄 Gerar Plano Alimentar", type="primary", disabled=not ready):
        if not os.path.exists(template_path):
            st.error(f"Template não encontrado: `{template_path}`")
        else:
            try:
                pdf_bytes = generate_final_pdf(template_path, anamnese, opcoes)
                nome_pet  = anamnese.get("pet_nome", "pet").replace(" ", "_")
                filename  = f"Plano_{nome_pet}_{datetime.now().strftime('%d%m%Y')}.pdf"
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
