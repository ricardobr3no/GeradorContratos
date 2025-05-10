from datetime import datetime
from flask import Flask, render_template, request, send_file
from main import converte_palavra_genero, criar_documento, salvar_documento, meses, limpar_console


app = Flask(__name__)

referencias = {}

@app.route("/")
def index():
    return render_template("formulario_contrato.html", context=referencias)


@app.route('/enviar', methods=['POST'])
def enviar():
    global referencias
    limpar_console()
    # print(list(request.form.items()))

    # criar um novo dicionario de referencias que será usado para fazer substituição no template docx
    referencias = {}
    # inicializa dicionario
    for name, value in request.form.items():
        # guarda nomes em MAIUSCULO
        if name.upper() in ('LOCADOR_NOME', 'LOCATARIO_NOME'): 
            value = value.upper()

        # adiciona value no dicionario de referencias
        referencias[name.upper()] = value


    # altera infos para masculino ou feminino
    for person in ('LOCADOR', 'LOCATARIO'):
        sexo = referencias[person + '_SEXO']

        estado_civil = referencias[person + '_ESTADO_CIVIL']
        referencias[person + '_ESTADO_CIVIL'] = converte_palavra_genero(estado_civil, sexo)
        

    # adicionar data e hora atual as referencias do contrato
    referencias["DD"] = str(datetime.now().day)
    referencias["MM"] = meses[datetime.now().month]
    referencias["YYYY"] = str(datetime.now().year)

    # substitui no template...
    novo_documento = criar_documento('template_contrato_de_aluguel_imovel.docx', referencias)
    file_path = salvar_documento(novo_documento, "contrato de aluguel - fulano.docx")
    return send_file(file_path, as_attachment=True)


if __name__ == '__main__':
    app.run(debug=True, port=5000)
