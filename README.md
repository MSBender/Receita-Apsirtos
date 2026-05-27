# Gerador de Documentos Veterinários — Apsirtos

App web para automatizar a geração de documentos da Dra. Isabelle Rizzo Assumpção (CRMV 48652/SP).

**Site:** https://receita-apsirtus.streamlit.app  
**Repositório:** https://github.com/MSBender/Receita-Apsirtos

---

## O que faz

A Isabelle acessa pelo navegador, faz upload da ficha cadastral do paciente e gera PDFs prontos para envio ao tutor — sem copiar e colar manualmente.

**4 tipos de documento:**
- Plano alimentar (dieta caseira)
- Plano alimentar (ração)
- Solicitação de exames
- Prescrição de medicamentos

---

## Como usar

1. Acesse `receita-apsirtus.streamlit.app`
2. Faça upload do `ANAMNESE.docx` — os dados do paciente e tutor são preenchidos automaticamente
3. Escolha o tipo de documento
4. Para dieta: faça upload dos prints do Dietalabs (o app lê os ingredientes automaticamente via IA)
5. Clique em **Gerar PDF** e baixe o arquivo pronto

---

## Stack

| Componente | Tecnologia |
|---|---|
| Interface | Python + Streamlit |
| Templates | python-docx (.docx) |
| Conversão PDF | LibreOffice headless |
| OCR Dietalabs | Claude Vision API (Haiku) |
| Merge de PDFs | pypdf |

---

## Estrutura de arquivos

```
receita-isabelle/
├── app.py                    # Interface Streamlit + lógica principal
├── fill_template.py          # Preenchimento dos templates Word
├── assets/
│   └── assinatura.png        # Assinatura digital (710×200px, PNG)
├── templates/
│   ├── template_dieta_caseira.docx
│   ├── template_dieta_racao.docx
│   ├── template_prescricao.docx
│   └── template_exames.docx
├── requirements.txt
├── packages.txt              # LibreOffice (para Streamlit Cloud)
├── iniciar.bat               # Rodar localmente no Windows
└── instalar.bat              # Instalar dependências localmente
```

---

## Rodapé dos documentos

Todos os documentos têm rodapé com:
- Assinatura digital da Isabelle (acima da linha separadora)
- Nome completo + CRMV (esquerda)
- Médica Veterinária (esquerda)
- Data de emissão (direita, alinhada à base)

**Configuração dos templates:** `bottom_margin = 3.5 cm`, `footer_distance = 0.8 cm`

---

## Deploy

O deploy é automático via GitHub → Streamlit Cloud. Após qualquer mudança:

```powershell
cd "C:\Users\KABUM\Documents\Claude\Projects\AUTOMATIZAÇÃO RECEITA ISABELLE\receita-isabelle"
git add .
git commit -m "descrição da mudança"
git push
```

O site atualiza em ~1 minuto após o push.

**Se aparecer erro de lock no git:**
```powershell
Remove-Item .git\index.lock -Force
Remove-Item .git\HEAD.lock -Force
```

---

## Variáveis de ambiente necessárias

| Variável | Onde configurar |
|---|---|
| `ANTHROPIC_API_KEY` | Windows: variável de sistema / Streamlit Cloud: Secrets |

---

## Campos da ficha ANAMNESE.docx

| Campo interno | Label no Word |
|---|---|
| `tutor_nome` | Nome completo / Nome do tutor |
| `tutor_cpf` | CPF |
| `tutor_email` | E-mail |
| `tutor_endereco` | Endereço |
| `pet_nome` | Nome do pet |
| `pet_especie` | Espécie |
| `pet_raca` | Raça |
| `pet_sexo` | Sexo |
| `pet_nascimento` | Data de nascimento |

---

## Histórico de mudanças relevantes

### 2026-05-27 — Feedback da Isabelle (v1.1)
- Corrigido parse do ANAMNESE.docx (campos chegavam em branco)
- Typo "Apsirtus" → "Apsirtos" em todos os lugares
- Coluna de quantidade na tabela de ingredientes: alinhada à esquerda
- OCR do Dietalabs: quantidades por extenso ("50 gramas" em vez de "50")
- Assinatura digital adicionada no rodapé de todos os documentos

### 2026-05 — Versão inicial (v1.0)
- App Streamlit completo com 4 tipos de documento
- Integração com Claude Vision para OCR do Dietalabs
- Templates Word com rodapé padronizado
- Deploy no Streamlit Cloud com UptimeRobot
