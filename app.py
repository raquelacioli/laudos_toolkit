import sys
import tempfile
from pathlib import Path
from typing import List, Optional

import streamlit as st

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

import preencher_laudo as pl

st.set_page_config(page_title="Gerador de Laudos Periciais", page_icon="📄", layout="wide")

st.title("Gerador de Laudos Periciais")
st.write(
    "Suba o modelo do Word, o arquivo com os dados da perícia e as fotos. "
    "O sistema gera um documento .docx já preenchido e pronto para revisão."
)


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


with st.sidebar:
    st.header("Passo a passo")
    st.markdown(
        "1. Faça upload do modelo .docx\n"
        "2. Faça upload do arquivo .txt com os dados\n"
        "3. Adicione fotos de documentos e, se quiser, uma foto de rosto\n"
        "4. Clique em gerar"
    )

    st.header("Exemplo de conteúdo do arquivo .txt")
    st.code(
        """FULANO DE TAL\nIDADE: 55 ANOS\nESCOLARIDADE: ENSINO MÉDIO COMPLETO\nESTADO CIVIL: SOLTEIRO\nOCUPAÇÃO HABITUAL: PEDREIRO\nTEMPO DE AFASTAMENTO: HÁ 2 ANOS\n\nANAMNESE:\nRelata dor lombar crônica há anos.\n\nEXAME FÍSICO E MENTAL:\nApresenta limitação funcional.\n\nCONCLUSÃO:\nPelo exposto, conclui-se que...""",
        language="text",
    )

modelo_file = st.file_uploader("Modelo do laudo (.docx)", type=["docx"])
dados_file = st.file_uploader("Dados da perícia (.txt)", type=["txt"])
fotos_files = st.file_uploader(
    "Fotos de documentos para anexo",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True,
)
face_photo_file = st.file_uploader(
    "Foto de rosto (opcional)",
    type=["jpg", "jpeg", "png"],
)
output_name = st.text_input("Nome do arquivo final", value="laudo_preenchido.docx")

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
