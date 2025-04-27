import os
import docx
from docx import document
from rich.progress import track


meses = {
    1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
    5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
    9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Decembro"
}


def criar_documento(template_path:str, referencias: dict) -> document.Document:
    documento = docx.Document(template_path)
    # modifica template
    for paragafo in track(documento.paragraphs, description="Preechendo contrato..."):
        for codigo in referencias:
            paragafo.text = paragafo.text.replace(codigo, referencias[codigo])
    return documento


def salvar_documento(documento: document.Document, nome_arquivo: str) -> str:
    DIR_NAME = 'contratos'  # diretorio em que serao salvos os contratos criados

    if not nome_arquivo.endswith((".docx", ".doc")):
        nome_arquivo += ".docx"

    if not os.path.exists(DIR_NAME):
        os.makedirs(DIR_NAME)
        print("diretorio 'contratos' criado!")

    FULLPATH = DIR_NAME + '/' + nome_arquivo
    documento.save(FULLPATH)
    return FULLPATH
