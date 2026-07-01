#!/usr/bin/env python3
"""
processar_lote.py
-------------------
Processa varios laudos periciais de uma vez, a partir de uma pasta com
uma subpasta para cada caso.

ESTRUTURA ESPERADA:

    casos/
        VALDETE_BEZERRA/
            modelo.docx          <- o modelo do laudo (qualquer nome .docx)
            dados.txt            <- os dados da pericia (qualquer nome .txt)
            rosto.jpg            <- (opcional) foto de rosto -- o nome do
                                     arquivo precisa CONTER "rosto" ou "face"
            doc1.jpg, doc2.jpg   <- (opcional) fotos de documentos p/ anexo
                                     (qualquer outro nome de imagem)
        VICENCIA_SANTANA/
            ...
        OUTRO_CASO/
            ...

Cada subpasta de "casos/" e tratada como um caso independente. Dentro dela:
- o unico arquivo .docx e usado como modelo;
- o unico arquivo .txt e usado como dados;
- imagens (.jpg/.jpeg/.png) cujo nome contenha "rosto" ou "face" viram a
  foto de rosto (3x4cm, pagina 1);
- as demais imagens viram o anexo de documentos (lado a lado, ao final).

USO:

    python3 processar_lote.py casos/ saida/

Gera, em saida/, um arquivo "<NOME_DA_PASTA>_PREENCHIDO.docx" para cada
caso processado com sucesso, alem de um resumo no final (quantos deram
certo, quais falharam e por que).
"""

import sys
import argparse
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import preencher_laudo as pl


IMAGE_EXTS = {".jpg", ".jpeg", ".png"}
FACE_HINTS = ("rosto", "face")


def find_case_files(case_dir: Path):
    docx_files = list(case_dir.glob("*.docx"))
    txt_files = list(case_dir.glob("*.txt"))
    images = [
        f for f in case_dir.iterdir()
        if f.is_file() and f.suffix.lower() in IMAGE_EXTS
    ]

    if len(docx_files) != 1:
        raise ValueError(
            f"esperado exatamente 1 arquivo .docx na pasta, encontrado {len(docx_files)}"
        )
    if len(txt_files) != 1:
        raise ValueError(
            f"esperado exatamente 1 arquivo .txt na pasta, encontrado {len(txt_files)}"
        )

    foto_rosto = None
    fotos_anexo = []
    for img in sorted(images):
        name_lower = img.stem.lower()
        if foto_rosto is None and any(h in name_lower for h in FACE_HINTS):
            foto_rosto = img
        else:
            fotos_anexo.append(img)

    return docx_files[0], txt_files[0], foto_rosto, fotos_anexo


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pasta_casos", help="Pasta contendo uma subpasta por caso")
    ap.add_argument("pasta_saida", help="Pasta onde salvar os laudos preenchidos")
    args = ap.parse_args()

    casos_dir = Path(args.pasta_casos)
    saida_dir = Path(args.pasta_saida)
    saida_dir.mkdir(parents=True, exist_ok=True)

    subpastas = sorted([p for p in casos_dir.iterdir() if p.is_dir()])
    if not subpastas:
        print(f"Nenhuma subpasta encontrada em {casos_dir}")
        return

    sucesso, falha = [], []

    for case_dir in subpastas:
        nome_caso = case_dir.name
        print(f"\n{'=' * 60}")
        print(f"Processando: {nome_caso}")
        print("=" * 60)
        try:
            modelo, dados_txt, foto_rosto, fotos_anexo = find_case_files(case_dir)
            output_path = saida_dir / f"{nome_caso}_PREENCHIDO.docx"
            work_dir = Path("/home/claude/work/_tmp_lote") / nome_caso

            log = pl.process_case(
                modelo=str(modelo),
                dados_txt=str(dados_txt),
                fotos=[str(f) for f in fotos_anexo],
                output=str(output_path),
                foto_rosto=str(foto_rosto) if foto_rosto else None,
                work_dir=str(work_dir),
            )
            for line in log:
                print(" ", line)
            print(f"OK -> {output_path}")
            sucesso.append(nome_caso)
        except Exception as e:
            print(f"FALHOU: {e}")
            falha.append((nome_caso, str(e)))

    print(f"\n\n{'#' * 60}")
    print("RESUMO DO LOTE")
    print("#" * 60)
    print(f"Processados com sucesso: {len(sucesso)}")
    for n in sucesso:
        print(f"  OK   {n}")
    print(f"Falharam: {len(falha)}")
    for n, err in falha:
        print(f"  FALHOU  {n}: {err}")


if __name__ == "__main__":
    main()
