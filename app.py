from flask import Flask, render_template, request
from openpyxl import Workbook

app = Flask(__name__)


# Rota para a página inicial
@app.route('/')
def index():
    return render_template('index.html')


# Rota para lidar com o envio do formulário
@app.route('/submit_form', methods=['POST'])
def submit_form():
    if request.method == 'POST':
        nome = request.form['nome']
        email = request.form['email']

        # Adiciona os dados a uma planilha do LibreOffice Calc
        adicionar_dados_planilha(nome, email)

        return 'Dados enviados com sucesso!'


def adicionar_dados_planilha(nome, email):
    # Cria uma nova planilha e adiciona os dados
    planilha = Workbook()
    planilha_ativa = planilha.active
    planilha_ativa.append(['Nome', 'Email'])
    planilha_ativa.append([nome, email])

    # Salva a planilha
    planilha.save('dados_clientes.xlsx')


if __name__ == '__main__':
    app.run(debug=True)
