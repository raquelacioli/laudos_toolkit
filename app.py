import sys
import tempfile
from pathlib import Path
from typing import List, Optional

import streamlit as st

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

import preencher_laudo as pl

st.set_page_config(page_title="Gerador de Laudos Periciais", page_icon="📄", layout="wide")

st.markdown(
    "<style>"
    "body {background-color: #f3f5f8;}"
    ".stApp {color: #1f2937; font-family: 'Segoe UI', sans-serif;}"
    ".login-card {background: #ffffff; padding: 36px 28px; border-radius: 22px; "
    "box-shadow: 0 20px 50px rgba(15, 23, 42, 0.08); max-width: 650px; margin: auto;}"
    ".login-card h1 {margin-bottom: 8px; color: #0b4d78;}"
    ".login-card p {color: #475569; margin-bottom: 24px; line-height: 1.6;}"
    ".info-box {background: #eef4fb; padding: 18px 20px; border-radius: 16px; margin-bottom: 20px;}"
    ".app-header {background: linear-gradient(135deg, #0b4d78 0%, #1a8cd8 100%); padding: 28px; border-radius: 24px; color: white; margin-bottom: 24px;}"
    ".app-header h1 {margin: 0; font-size: 2.3rem;}"
    ".app-header p {margin: 8px 0 0; color: rgba(255,255,255,0.84); font-size: 1rem;}"
    "</style>",
    unsafe_allow_html=True,
)

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

if not st.session_state.logged_in:
    st.markdown(
        "<div class='login-card'>"
        "<h1>Login seguro</h1>"
        "<p>Digite seu e-mail e senha para acessar a plataforma de perícias médicas.</p>"
        "</div>",
        unsafe_allow_html=True,
    )
    st.write(" ")
    login_col1, login_col2, login_col3 = st.columns([1, 2, 1])
    with login_col2:
        email = st.text_input("E-mail", key="login_email")
        password = st.text_input("Senha", type="password", key="login_password")
        if st.button("Entrar", type="primary"):
            if email == "tiagoacioli@gmail.com" and password == "sucesso":
                st.session_state.logged_in = True
                if hasattr(st, "rerun"):
                    st.rerun()
                else:
                    st.experimental_rerun()
            else:
                st.error("Login ou senha incorretos.")
    st.stop()

st.markdown(
    "<div class='app-header'>"
    "<h1>Gerador de Laudos Periciais</h1>"
    "<p>Plataforma profissional para gerar laudos a partir de modelos, dados clínicos e anexos.</p>"
    "</div>",
    unsafe_allow_html=True,
)

st.write("### Painel de entrada")
st.info("Envie os arquivos necessários e gere o laudo automaticamente. O resultado fica disponível para download imediato.")


def generate_docx_from_uploads(
    modelo_path: Path,
    dados_txt_path: Path,
    fotos_paths: List[Path],
    output_path: Path,
    foto_rosto_path: Optional[Path] = None,
    work_dir: Optional[Path] = None,
):
    work = work_dir or output_path.parent / "work"
    return pl.process_case(
        modelo=str(modelo_path),
        dados_txt=str(dados_txt_path),
        fotos=[str(p) for p in fotos_paths],
        output=str(output_path),
        foto_rosto=str(foto_rosto_path) if foto_rosto_path else None,
        work_dir=str(work),
    )


col1, col2 = st.columns([3, 1])
with col2:
    st.markdown(
        "<div style='background:#f8f9fb; padding:18px; border-radius:14px;'>"
        "<h3 style='margin:0;'>Usuário logado</h3>"
        "<p style='margin:6px 0 12px 0;'>tiagoacioli@gmail.com</p>"
        "</div>",
        unsafe_allow_html=True,
    )
    if st.button("Logout", key="logout_button"):
        st.session_state.logged_in = False
        st.experimental_rerun()

with st.sidebar:
    st.markdown(
        "<div style='background:#ffffff; padding:20px; border-radius:18px; box-shadow: 0 20px 40px rgba(15,23,42,0.08);'>"
        "<h2 style='margin-top:0; color:#0b4d78;'>Instruções</h2>"
        "<p style='color:#475569; line-height:1.75;'>Siga os passos abaixo para gerar um laudo pericial completo e seguro.</p>"
        "</div>",
        unsafe_allow_html=True,
    )
    st.markdown("### Passo a passo")
    st.write("1. Faça upload do modelo .docx")
    st.write("2. Faça upload do arquivo .txt com os dados")
    st.write("3. Adicione fotos de documentos e, se quiser, uma foto de rosto")
    st.write("4. Clique em gerar")
    st.markdown("---")
    st.markdown("**Recomendações:**")
    st.write("- Use modelo com campos vazios e tabelas limpas.")
    st.write("- Dados em .txt devem estar completos e sem erros de digitação.")
    st.write("- Fotos claras melhoram a qualidade do anexo final.")

with st.container():
    st.markdown(
        "<div style='background: #ffffff; padding: 30px; border-radius: 24px; "
        "box-shadow: 0 24px 60px rgba(15, 23, 42, 0.08); margin-bottom: 20px;'>"
        "<h2 style='margin-top:0; color:#0b4d78;'>Envio de arquivos</h2>"
        "<p style='color:#475569; line-height:1.7;'>Suba o modelo, os dados e as fotos para gerar o laudo. A saída é um documento Word pronto para revisão.</p>"
        "</div>",
        unsafe_allow_html=True,
    )
    modelo_file = st.file_uploader("Modelo do laudo (.docx)", type=["docx"], label_visibility="visible")
    dados_file = st.file_uploader("Dados da perícia (.txt)", type=["txt"], label_visibility="visible")
    fotos_files = st.file_uploader(
        "Fotos de documentos para anexo",
        type=["jpg", "jpeg", "png"],
        accept_multiple_files=True,
        label_visibility="visible",
    )
    face_photo_file = st.file_uploader(
        "Foto de rosto (opcional)",
        type=["jpg", "jpeg", "png"],
        label_visibility="visible",
    )
    output_name = st.text_input("Nome do arquivo final", value="laudo_preenchido.docx")

    st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)
    if st.button("Gerar laudo", type="primary"):
        if not modelo_file or not dados_file:
            st.error("Envie pelo menos o modelo .docx e o arquivo .txt antes de gerar.")
            st.stop()

        with st.spinner("Processando o laudo e anexos..."):
            tmp_dir = Path(tempfile.mkdtemp(prefix="laudo_streamlit_"))
            try:
                modelo_path = tmp_dir / modelo_file.name
                dados_path = tmp_dir / dados_file.name
                modelo_path.write_bytes(modelo_file.getvalue())
                dados_path.write_text(dados_file.getvalue().decode("utf-8"), encoding="utf-8")

                fotos_paths = []
                for uploaded in fotos_files or []:
                    photo_path = tmp_dir / uploaded.name
                    photo_path.write_bytes(uploaded.getvalue())
                    fotos_paths.append(photo_path)

                face_path = None
                if face_photo_file is not None:
                    face_path = tmp_dir / face_photo_file.name
                    face_path.write_bytes(face_photo_file.getvalue())

                output_path = tmp_dir / output_name
                if not output_name.lower().endswith(".docx"):
                    output_path = tmp_dir / f"{output_name}.docx"

                log = generate_docx_from_uploads(
                    modelo_path=modelo_path,
                    dados_txt_path=dados_path,
                    fotos_paths=fotos_paths,
                    output_path=output_path,
                    foto_rosto_path=face_path,
                    work_dir=tmp_dir / "work",
                )

                if not output_path.exists():
                    raise RuntimeError("O arquivo final não foi criado.")

                bytes_data = output_path.read_bytes()
                st.success("Laudo gerado com sucesso.")
                st.download_button(
                    label="Baixar documento .docx",
                    data=bytes_data,
                    file_name=output_path.name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )

                st.subheader("Log do processamento")
                for line in log:
                    st.write(line)
            finally:
                import shutil

                shutil.rmtree(tmp_dir, ignore_errors=True)
