from datetime import datetime
from flask import Flask, render_template, request, send_file, send_from_directory
from main import criar_documento, salvar_documento, meses
import os

limpar_console = lambda : os.system('clear' if os.name != "nt" else "clear")

app = Flask(__name__)


@app.route("/")
def index():
    return render_template("formulario_contrato.html")


@app.route('/enviar', methods=['POST'])
def enviar():
    limpar_console()
    # print(list(request.form.items()))

    # criar um dicionario de referencias que será usado para fazer substituição no template docx
    referencias = {}
    for name, value in request.form.items():
        referencias[name.upper()] = value
    # adicionar data e hora atual as referencias do contrato
    referencias["DD"] = str(datetime.now().day)
    referencias["MM"] = meses[datetime.now().month]
    referencias["YYYY"] = str(datetime.now().year)
    # print(referencias)

    # substitui no template...
    novo_documento = criar_documento('template_contrato_de_aluguel_imovel.docx', referencias)
    file_path = salvar_documento(novo_documento, "contrato de aluguel - fulano.docx")
    return send_file(file_path, as_attachment=True)


if __name__ == '__main__':
    app.run(debug=True, port=5000)
