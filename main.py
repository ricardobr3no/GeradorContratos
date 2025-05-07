import os
from datetime import datetime

from docx import Document
from docx.document import Document as Document_type
from docx.text.paragraph import Paragraph
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from rich.progress import track


limpar_console = lambda : os.system('clear' if os.name != "nt" else "clear")

meses = {
    1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
    5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
    9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Decembro"
}

def formata_paragrafo(documento: Document_type) -> None:
    # adiciona negrito a primeira palavra quando tiver classula
    p_data =  f"São Luís, {str(datetime.now().day).zfill(2)} de {meses[datetime.now().month]} de {datetime.year}."
    in_ceter = False

    for i, paragafo in enumerate(documento.paragraphs):
        if i == 0: # primeira paragafo
            paragafo.alignment =  WD_PARAGRAPH_ALIGNMENT.CENTER  # paragafo justificado
            text_run = paragafo.runs[0]
            text_run.underline = True
            text_run.bold = True
            paragafo.runs[0] = text_run

        elif paragafo.runs:
            first_run = paragafo.runs[0]
            
            if paragafo.text == p_data:
                in_ceter = True
            if in_ceter:
                paragafo.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            if paragafo.text.startswith('Testemunha'):
                in_ceter = False

            if first_run.text:
                words = first_run.text.split(sep=":", maxsplit=1)
                first_word = words[0] + ":"if len(words) > 1 else words[0]
                remaing_text = words[1] if len(words) > 1 else ""
                first_run.bold = True if len(words) > 1 else False
                first_run.text = ''
                first_run.add_text(first_word)
                paragafo.add_run(remaing_text)
                paragafo.runs[0] = first_run

            print(first_run.text)



def criar_documento(template_path:str, referencias: dict) -> Document_type:
    documento = Document(template_path)
    # modifica template
    for paragafo in track(documento.paragraphs, description="Preechendo contrato..."):
        paragafo.alignment =  WD_PARAGRAPH_ALIGNMENT.JUSTIFY  # paragafo justificado

        for codigo in referencias:
            paragafo.text = paragafo.text.replace(codigo, referencias[codigo])


    formata_paragrafo(documento)
    return documento


def salvar_documento(documento: Document_type, nome_arquivo: str) -> str:
    DIR_NAME = 'contratos'  # diretorio em que serao salvos os contratos criados

    if not nome_arquivo.endswith((".docx", ".doc")):
        nome_arquivo += ".docx"

    if not os.path.exists(DIR_NAME):
        os.makedirs(DIR_NAME)
        print("diretorio 'contratos' criado!")

    FULLPATH = DIR_NAME + '/' + nome_arquivo
    documento.save(FULLPATH)
    return FULLPATH
