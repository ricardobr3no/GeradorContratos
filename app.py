from datetime import datetime
from flask import Flask, render_template, request, send_file 
from main import (converte_palavra_genero, criar_documento,
                    salvar_documento, meses, limpar_console)
import requests


app = Flask(__name__)
referencias = {}


def array_uf():
    response = requests.get('https://servicodados.ibge.gov.br/api/v1/localidades/estados')
    array = []
    if response.ok:
        for UF in response.json():
            array.append(UF['sigla'])
    return array


def city_uf(uf: str):
    response = requests.get(f'https://servicodados.ibge.gov.br/api/v1/localidades/estados/{uf.upper()}/municipios')
    cidades = []
    if response.ok:
        for cidade in response.json():
            cidades.append(cidade['nome'])
    return cidades


def update_fields():
    global referencias
    # criar um novo dicionario de referencias que será usado para fazer substituição no template docx
    referencias = {}
    # inicializa dicionario
    for name, value in request.form.items():
        print(name, value)
        # guarda nomes em MAIUSCULO
        if name.upper() in ('LOCADOR_NOME', 'LOCATARIO_NOME'): 
            value = value.upper()

        # adiciona value no dicionario de referencias
        referencias[name.upper()] = value

    # print('atualizou'.upper())



@app.route("/")
def index():
    locador_uf = request.args.get('locador_uf')
    imovel_uf = request.args.get('imovel_uf')

    return render_template('formulario_contrato.html',
                           array_uf=array_uf(), 
                           locador_uf=locador_uf if locador_uf else 'MA',
                           imovel_uf=imovel_uf if imovel_uf else 'MA', 
                           city_locador_uf=city_uf(locador_uf if locador_uf else 'MA'), 
                           city_imovel_uf=city_uf(imovel_uf if imovel_uf else 'MA'), 
                           context=referencias)


@app.route('/enviar', methods=['POST'])
def enviar():
    global referencias
    limpar_console()
    # print(list(request.form.items()))
    update_fields()
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
