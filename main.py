from flask import Flask

app = Flask(__name__)
app.secret_key = 'troque-esta-chave-em-producao'
app.config['MAX_CONTENT_LENGTH'] = 5 * 1024 * 1024  # limite de 5MB para upload de planilhas

from routes import *

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
