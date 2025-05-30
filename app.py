#Importando os módulos necessários para o Flask funcionar
from flask import Flask, render_template,request,redirect
#Importando as funções da biblioteca openpyxl para criar e
#manipular um arquivo excel
from openpyxl import Workbook, load_workbook
#Biblioteca para verificar a existência de um arq. excel
import os
#Criar de fato a nossa aplicação
app = Flask(__name__)
#Definindo o nome da planilha excel
ARQUIVO = 'vendas.xlsx'

if not os.path.exists(ARQUIVO):
    wb = Workbook()#CRIANDO UM ARQ. EXCEL
    ws = wb.active #Selecionando um plano. ativa do projeto
#Criando um cabeçalho para a planilha
    ws.append(["Nome","Vendas","Metas"])
#salvando o arquivo
    wb.save(ARQUIVO)

#22/05
#Rota principal do site (formulário cadastro de vendas)
@app.route('/')
def index():
    return render_template('index.html')
#função que deve ser executada, ao ser requisitado a rota '/'
#abrir a página principal, neste caso é o index.html

#Rota que processa os dados do formulário e salva no Excel
@app.route('/salvar',methods=["POST"])
def salvar():
#CAPTURA OS DADOS DE CADA UMA DAS CAIXAS DO FORMULÁRIO E ATRIBUI PARA
#AS VARIÁVEIS
    nome = request.form['nome']
#CAPTURA A INFO. DA CX. FAÇA O FORMULÁRIO. CONVERTE PARA FLOAT E ATRIBUI À
#VARIÁVEL
    vendas = float(request.form['vendas'])
    meta = float(request.form['meta'])
#ABRINDO O ARQUIVO EXCEL
    wb = load_workbook(ARQUIVO)
#SELECIONANDO A PLANILHA ATIVA - 1ª ABA Por padrão
    ws = wb.active

#Adiciona uma nova linha como lista com as informações do formulário
    ws.append([nome,vendas,meta])

#Salvando o arquivo excel
    wb.save(ARQUIVO)

#Redirecionando a rota para /analisar (onde abrirá uma nova página)
#Passando por parâmetro o nome do funcionário.
    return redirect('/analisar?nome='+nome)
#Rota para a tela resultado... analisando se o funcionário bateu a meta
@app.route('/analisar')
def analisar():
#Pegar o parâmetro do nome da func. enviado como parâmetro para url
    nome_param = request.args.get('nome')

    wb = load_workbook(ARQUIVO)
    ws = wb.active

#Loop for: percorre todas as linhas do plano. a partir da 2ª linha
#(Pois a 1ª linha é cabeçalho) valores_only retorna apenas os valores
#da linha em
    for linha in ws.iter_rows(min_row=2, values_only=True):
        nome, vendas, meta = linha
    # a linha sempre recebe três valores, ex:
    # linha = ('Ana', 45, 50) na linha de código acima, estamos
    # atribuindo cada elemento da linha a uma variável na sequencia
    # as variáveis ​​nome, vendas e meta recebem Ana, 45, 50
    # respectivamente. Nome técnico desse processo é desempacotamento

    #Verifique se o nome atual da linha é o mesmo enviado na URL
        if nome == nome_param:
            meta_batida = (vendas >= meta)
        #Se a meta for batida bônus recebe o resultado do cálculo de
        # 15% do valor da venda, caso contrário receberá 0
            bonus = round(vendas * 0.15, 2) if meta_batida else 0
        #Exibir a tela resultante com as informações dos cálculos
            return render_template('resultado.html',
            nome = nome,meta_batida = meta_batida,
            bonus = bonus)
    return "Funcionário não encontrado"

#Rota que mostra a pág. do histórico de todos os funcionários cadastrados
@app.route('/historico')
def historico():
    wb = load_workbook(ARQUIVO)
    ws = wb.active
#Converte os dados da planilha ( a partir da 2ª linha ) em uma tupla
    dados = list(ws.iter_rows(min_row = 2,values_only=True))
    return render_template('historico.html', dados = dados)
    #Iniciando o Flask no modo desenvolvedor Debug
if __name__ == '__main__':
        app.run(debug=True)