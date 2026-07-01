# Como usar o preencher_laudo.py

## O que ele faz
1. Lê os dados de um arquivo `.txt` (formato livre, explicado abaixo);
2. Preenche automaticamente os campos do modelo `.docx` do laudo (incluindo os
   campos de menu suspenso do Word — Escolaridade, Estado civil, Ocupação
   habitual, etc. — escolhendo a opção mais parecida com o que você escreveu);
3. Anexa, ao final do documento, as fotos de documentos como folha de anexo.

## Comando

```bash
python3 preencher_laudo.py MODELO.docx DADOS.txt foto1.jpg foto2.jpg --foto-rosto rosto.jpg -o SAIDA.docx
```

- `MODELO.docx`: o modelo do laudo (com os campos em branco/menu suspenso);
- `DADOS.txt`: arquivo de texto com os dados da perícia (veja formato abaixo);
- `foto1.jpg foto2.jpg ...`: fotos de documentos para o anexo, na ordem em
  que devem aparecer. Cada foto passa por um processamento automático (tipo
  scanner) antes de entrar no laudo: o script detecta o documento na foto,
  recorta o fundo (mesa, sombra, etc.), corrige rotação/perspectiva e
  melhora nitidez e contraste. Depois, encaixa cada documento numa caixa de
  até 10x5,5cm **preservando a proporção original** (sem esticar/distorcer)
  — por isso documentos na horizontal e na vertical ficam com tamanhos de
  exibição diferentes, lado a lado, 2 por linha. Como isso não cabe numa
  página vertical comum, essas páginas do anexo saem em **paisagem**
  automaticamente; o resto do laudo continua em retrato normalmente. Requer
  o arquivo `scan_documento.py` na mesma pasta;
- `-o SAIDA.docx`: nome do arquivo final;
- `--foto-rosto FOTO.jpg`: (opcional) foto de rosto do(a) periciado(a), 3x4cm.
  O script procura a coluna reservada ao lado do rótulo "DADOS DO(A)
  PERICIADO(A)", na tabela do preâmbulo. **Se essa coluna estiver muito
  estreita** (menos de ~1,5cm), o script avisa no log e não insere a foto —
  nesse caso, basta abrir o modelo no Word e alargar a coluna (arrastando a
  borda) para abrir espaço.

## Formato do arquivo .txt

Pode escrever como você já escreve naturalmente. Regras:

- A primeira linha solta (sem `CAMPO:`) é interpretada como o nome do(a)
  periciado(a) — hoje ele só é usado para referência, pois o modelo não tem
  campo de nome no corpo do laudo;
- Campos simples: `CAMPO: valor` na mesma linha
  (ex: `IDADE: 61 ANOS`, `ESTADO CIVIL: CASADA`);
- Campos de texto longo (anamnese, exame físico, conclusão): coloque o
  cabeçalho sozinho na linha (com ou sem `:`) e o texto nas linhas seguintes,
  até a próxima linha de cabeçalho;
- Pequenos erros de digitação nos cabeçalhos são tolerados
  (ex: "ANAMENESE", "CONCLUISÃO" são reconhecidos).

Campos reconhecidos hoje: IDADE, ESCOLARIDADE, ESTADO CIVIL,
OCUPAÇÃO HABITUAL, TEMPO DE AFASTAMENTO, SEXO, RG, CPF, ANAMNESE,
EXAME FÍSICO E MENTAL, CONCLUSÃO.

Exemplo:

```
FULANO DE TAL
IDADE: 55 ANOS
ESCOLARIDADE: ENSINO FUNDAMENTAL INCOMPLETO
ESTADO CIVIL: DIVORCIADO
OCUPAÇÃO HABITUAL: PEDREIRO
TEMPO DE AFASTAMENTO: HÁ 2 ANOS

ANAMNESE:
Relata dor lombar crônica há cerca de 3 anos, com irradiação para o
membro inferior direito...

EXAME FÍSICO E MENTAL:
Apresenta limitação de flexão do tronco...

CONCLUSÃO:
Pelo exposto, conclui-se que...
```

## Processando vários casos de uma vez (lote)

Para processar centenas de casos sem digitar o comando um por um, use o
`processar_lote.py`. Ele espera uma pasta com **uma subpasta para cada
caso**:

```
casos/
    VALDETE_BEZERRA/
        modelo.docx          <- o modelo do laudo (qualquer nome .docx)
        dados.txt            <- os dados da perícia (qualquer nome .txt)
        rosto.jpg            <- (opcional) foto de rosto — o NOME do
                                 arquivo precisa conter "rosto" ou "face"
        doc1.jpg, doc2.jpg   <- (opcional) fotos de documentos p/ anexo
                                 (qualquer outro nome de imagem)
    VICENCIA_SANTANA/
        ...
    OUTRO_CASO/
        ...
```

Regras de identificação automática dentro de cada subpasta:
- o único arquivo `.docx` é o modelo;
- o único arquivo `.txt` são os dados;
- imagens cujo nome contenha "rosto" ou "face" viram a foto de rosto;
- as demais imagens (jpg/jpeg/png) viram o anexo de documentos.

Rodando:

```bash
python3 processar_lote.py casos/ saida/
```

Gera, em `saida/`, um arquivo `NOME_DA_PASTA_PREENCHIDO.docx` para cada
caso. No final, mostra um resumo (quantos deram certo, quais falharam e
por quê) — **um caso com problema não trava os outros**. Requer
`preencher_laudo.py` e `scan_documento.py` na mesma pasta.


Hoje o script não mexe em: RG/CPF (só se você informar), datas
(nascimento, requerimento, perícia), CID/patologias (tabela de
diagnóstico), e os quesitos. Dá pra estender o script para cobrir isso
também — é só me avisar quais campos você quer automatizar em seguida.

## Ao final, sempre confira o resultado
O script mostra um log de quais campos foram encontrados e preenchidos
(`OK`) e quais não foram localizados no modelo (`--`). Sempre abra o
.docx gerado e revise antes de assinar/protocolar — a automação escolhe
a opção mais parecida no menu suspenso, mas pode errar em casos
ambíguos.
