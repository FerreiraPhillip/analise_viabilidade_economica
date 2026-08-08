from main import app
from flask import render_template

@app.route('/')
def homepage():
    return render_template('homepage.html')

@app.route('/Analise de viabilidade')
def analise_viabilidade():
    return render_template('analise_viabilidade.html')