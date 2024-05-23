from flask import Flask, request, json, Response
from pymongo import MongoClient
import pymongo
import logging as log
import win32com.client as win32
from flask import Flask, request, jsonify
from pymongo import MongoClient
import datetime
import pythoncom


#função para mandar email 

def email(email,payer,payee,valor):
    try:
        pythoncom.CoInitialize()
        outlook = win32.Dispatch('Outlook.Application')
        email_out = outlook.CreateItem(0)

        texto=f"""
        <title>Resumo da Transferência</title>
        <body>
        <h1>Resumo da Transferência</h1>
        <p><strong>Valor da Transferência:</strong>{valor}</p>
        <p><strong>Origem:</strong> {payer}</p>
        <p><strong>Destino:</strong> {payee}</p>"""

        email_out.To = email
        email_out.Subject = 'Resumo da Transferência'
        email_out.HTMLBody = texto

        email_out.Send()
    except Exception as e:
        return jsonify({f'erro':'erro ao enviar o email {e}'})



	

app = Flask(__name__)



#conexão com o mongo
client = pymongo.MongoClient("mongodb://localhost:27017/desafio")
db = client["desafio"]
collection = db["cliente"]                                                                 



def read(self):
    log.info('Reading All Data')
    documents = self.collection.find()
    output = [{item: data[item] for item in data if item != '_id'} for data in documents]
    return output



@app.route('/inserir', methods=['POST'])
def inserir():
    campos_obrigatorios = ['email', 'nome', 'senha', 'cpf','id']
    dados = request.json

    email = dados.get('email')
    cpf = dados.get('cpf')
    id = dados.get('id')
    usuario = dados.get('usuario')
    saldo = dados.get('saldo')
    senha = dados.get('senha')
    nome = dados.get('nome')

    cliente = {'email':email,
               'cpf':cpf,
               'id':id,
               'usuario':usuario,
               'saldo':saldo,
               'senha':senha,
               'nome':nome}
    

    #verificando os campos obrigatorios
    for campo in campos_obrigatorios:
        if campo not in dados:
            return jsonify({'erro': 'Há campos obrigatorios não preenchidos'}), 400

    #verificando se o cpf ou email ja esta no banco
    if collection.find_one({'email':email}):
        return jsonify({'erro': 'O cliente com esse email já esta cadastrado'}), 400

    if collection.find_one({'cpf':cpf}):
        return jsonify({'erro': 'O cliente com esse cpf já esta cadastrado '}), 400
    

    #verificando se id ja esta no banco
    if collection.find_one({'id':id}):
        return jsonify({'erro': 'O cliente com esse id já esta cadastrado '}), 400
    

    #verificando se é logista ou usuario
    if not usuario == 'comum' and not usuario == 'lojista':
        return jsonify({'erro': usuario}), 400

    #inserindo cliente no banco de dados
    collection.insert_one(cliente)
    return jsonify({'mensagem': 'Cliente inserido com sucesso'}), 201




@app.route('/transferir', methods=['PUT'])
def transferir():
    dados = request.json

    #parametros
    payer = dados.get('payer') #origem
    payee = dados.get('payee') #destino
    value = dados.get('value')

     #verificando se a conta origem é lojista
    if collection.find_one({'id': payer}).get('usuario') == 'lojista':
        return jsonify({'erro': 'Contas de lojistas não podem realizar transferencias'}), 400

    #verificando se a conta origem existe
    if not collection.find_one({'id': payer}):
        return jsonify({'erro': 'Conta de origem não encontrada'}), 400

    #verificando se a conta destino existe
    if not collection.find_one({'id': payee}):
        return jsonify({'erro': 'Conta de destino não encontrada'}), 400

    #verificando se há saldo na conta origem
    if collection.find_one({'id': payer})['saldo'] < value:
        return jsonify({'erro':'saldo insuficiente'}),400
    

    #dando update nos documentos

    emailr = collection.find_one({'id': payer}).get('email')
    emaile = collection.find_one({'id': payee}).get('email')
    nomer = collection.find_one({'id': payer}).get('nome')
    nomee = collection.find_one({'id': payee}).get('nome')
    print(nomer,nomee)

    try:
        collection.update_one({'id': payer}, {'$inc': {'saldo': -value}})
        collection.update_one({'id': payee}, {'$inc': {'saldo': value}})
        email(emailr,nomer,nomee,value)
        return jsonify({'mensagem':'transação realizada com sucesso'}), 201
        
    except Exception as e:
        collection.update_one({'id': payer}, {'$inc': {'saldo': value}})
        collection.update_one({'id': payee}, {'$inc': {'saldo': -value}})
        return jsonify({'erro':'erro ao transferir o dinheiro entre as contas'})
    
    
    
    




if __name__ == '__main__':
    app.run(debug=True, port=5001, host='0.0.0.0')
    