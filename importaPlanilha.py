import pandas as pd
import requests
import sys
import datetime
import mysql.connector
import os

mydb = mysql.connector.connect(
  host="<Database Server IP>",
  port="<mysql port>",
  user="<user_db>",
  password="<user_db_pass>",
  database="redmine_db"
)

fixversion_name = sys.argv[1]
executor_email = sys.argv[2]

####Dados email
para = "to: "+ executor_email
de = "From: devops@7comm.com.br"
assunto = "subject: Horas Importadas com Sucesso!"
assuntoerro = "subject: Falha ao importar horas!"
email = '<div style="width: 80%;display: flex; flex-direction: row;flex-wrap: wrap;justify-content: center;align-items: center;font-family: Arial; color: gray; background-color: white">   <div style=" width: 100%;height: 20%;margin: 10px;">      <div style = "width: 80%;height: auto;position: relative; padding: 10px;margin: 0 auto; border-bottom: 1px solid black">         <h1 style= "text-align:center; width:40%;margin: 0 auto";>Horas importadas com sucesso!</h1>      </div>   </div>   <div style="height: 100%; width:100%; margin: 0 auto;">      <p style = "margin:0 auto;width: 60%">Você está recebendo esse e-mail porque horas foram importadas com <span style = "color:green">sucesso</span> para o Redmine com via Jenkins</a>.		</p>         </div></div>'
###
mycursor = mydb.cursor()

if fixversion_name == "":
    print("Campo fixversion vazio")
    ###email de erro
    email = '<div style="width: 80%;display: flex; flex-direction: row;flex-wrap: wrap;justify-content: center;align-items: center;font-family: Arial; color: gray; background-color: white">   <div style=" width: 100%;height: 20%;margin: 10px;">      <div style = "width: 80%;height: auto;position: relative; padding: 10px;margin: 0 auto; border-bottom: 1px solid black">         <h1 style= "text-align:center; width:40%;margin: 0 auto";>Falha ao Importar horas!</h1>      </div>   </div>   <div style="height: 100%; width:100%; margin: 0 auto;">      <p style = "margin:0 auto;width: 60%">Você está recebendo esse e-mail porque houve uma <span style = "color:red">falha</span>  ao importar horas para o Redmine com via Jenkins</a>.</p> <br> <p style = "color:red">ERRO: Campo fixversion em branco</p>         </div></div>'
    sourceFile = open('email.txt', 'w')
    print(para, file = sourceFile)
    print(de, file = sourceFile)
    print(assuntoerro, file = sourceFile)
    print("MIME-Version: 1.0", file = sourceFile)
    print("content-type: text/html; charset=utf-8", file = sourceFile)
    print(email, file = sourceFile)
    sourceFile.close()
    os.system("ssmtp " + para + " < email.txt")
    ###fim email de erro
    
    sys.exit(1)

mycursor.execute("SELECT project_id FROM redmine_db.versions where name LIKE '%" + fixversion_name + "%' LIMIT 1;")
project_id = mycursor.fetchone()

if project_id is None:
    print("Fixversion não encontrada")
    ###email de erro
    email = '<div style="width: 80%;display: flex; flex-direction: row;flex-wrap: wrap;justify-content: center;align-items: center;font-family: Arial; color: gray; background-color: white">   <div style=" width: 100%;height: 20%;margin: 10px;">      <div style = "width: 80%;height: auto;position: relative; padding: 10px;margin: 0 auto; border-bottom: 1px solid black">         <h1 style= "text-align:center; width:40%;margin: 0 auto";>Falha ao Importar horas!</h1>      </div>   </div>   <div style="height: 100%; width:100%; margin: 0 auto;">      <p style = "margin:0 auto;width: 60%">Você está recebendo esse e-mail porque houve uma <span style = "color:red">falha</span>  ao importar horas para o Redmine com via Jenkins</a>.</p> <br> <p style = "color:red">ERRO: Fixversion não encontrada!</p>         </div></div>'
    sourceFile = open('email.txt', 'w')
    print(para, file = sourceFile)
    print(de, file = sourceFile)
    print(assuntoerro, file = sourceFile)
    print("MIME-Version: 1.0", file = sourceFile)
    print("content-type: text/html; charset=utf-8", file = sourceFile)
    print(email, file = sourceFile)
    sourceFile.close()
    os.system("ssmtp " + para + " < email.txt")
    ###fim email de erro

    sys.exit(1)

project_id=str(project_id[0])


mycursor.execute("SELECT tarefa_id FROM redmine_db.project_tarefabradesco WHERE project_id = "+ project_id +";")
tarefa_id = mycursor.fetchone()

if tarefa_id is None:
    print("Não existe uma tarefa para lançamento de horas bradesco cadastrada, crie a tarefa e/ou cadastre ela na base project_tarefabradesco (admin)!")
    #NÃO ESQUECER DE MONTAR O E-MAIL DE ERRO COM Acima
    ###email de erro
    email = '<div style="width: 80%;display: flex; flex-direction: row;flex-wrap: wrap;justify-content: center;align-items: center;font-family: Arial; color: gray; background-color: white">   <div style=" width: 100%;height: 20%;margin: 10px;">      <div style = "width: 80%;height: auto;position: relative; padding: 10px;margin: 0 auto; border-bottom: 1px solid black">         <h1 style= "text-align:center; width:40%;margin: 0 auto";>Falha ao Importar horas!</h1>      </div>   </div>   <div style="height: 100%; width:100%; margin: 0 auto;">      <p style = "margin:0 auto;width: 60%">Você está recebendo esse e-mail porque houve uma <span style = "color:red">falha</span>  ao importar horas para o Redmine com via Jenkins</a>.</p> <br> <p style = "color:red">ERRO: Não existe uma tarefa para lançamento de horas bradesco cadastrada, crie a tarefa e/ou cadastre ela na base project_tarefabradesco (admin)!</p>         </div></div>'
    sourceFile = open('email.txt', 'w')
    print(para, file = sourceFile)
    print(de, file = sourceFile)
    print(assuntoerro, file = sourceFile)
    print("MIME-Version: 1.0", file = sourceFile)
    print("content-type: text/html; charset=utf-8", file = sourceFile)
    print(email, file = sourceFile)
    sourceFile.close()
    os.system("ssmtp " + para + " < email.txt")
    ###fim email de erro
    sys.exit(1)

###Planilha de horas###
df = pd.read_excel('novo.xls')[['Questão-chave','Emissão de resumo','Horas', 'data de Trabalho','Nome de Usuário', 'Nome da atividade']]
qtdeLinhas = len(df.index) 
print("Quantidade de linhas:")
print(qtdeLinhas)
#qtdeLinhas = qtdeLinhas + 1

####API###
redmineurl = "<Redmine_URL>"
token = "<redmine_user_api_token>"
url = redmineurl + "time_entries.xml?key=" + token
headers = {'Content-Type': 'application/xml'}

i=0
while i != qtdeLinhas:
    print("valor de i:")
    print(i)
    #print(df.iloc[i])
    #print(df.loc[i, 'Nome de Usuário'])
    chaveM = df.loc[i, 'Nome de Usuário']
    descricao = df.loc[i, 'Emissão de resumo']
    data = df.loc[i, 'data de Trabalho']
    data = data.strftime("%Y-%m-%d")
    horas = df.loc[i, 'Horas']
    horas = str(horas)
    task_id = tarefa_id[0]
    task_id = str(task_id)
    pula = False    
    mycursor.execute("SELECT user_id FROM redmine_db.login_chavem WHERE chavem = '"+ chaveM +"';")
    user_id = mycursor.fetchone() 
    
    if user_id is None:
        pula = True    
    
    if pula == False:
        print(i)
        print(user_id)
        user_id=str(user_id[0])
    
        body ="<time_entry>\r\n<hours>"+ horas +"</hours>\r\n<comments>"+ descricao +"</comments>\r\n<spent_on>"+ data  +"</spent_on>\r\n<issue_id>"+ task_id +"</issue_id>\r\n<user_id>"+ user_id +"</user_id>\r\n<activity_id>9</activity_id>\r\n</time_entry>\r\n"

        requests.request("POST", url, headers=headers, data=body, verify=False)
        print(descricao + "Inserido ")
    
    i=i+1

print(email)    
sourceFile = open('email.txt', 'w')
print(para, file = sourceFile)
print(de, file = sourceFile)
print(assunto, file = sourceFile)
print("MIME-Version: 1.0", file = sourceFile)
print("content-type: text/html; charset=utf-8", file = sourceFile)
print(email, file = sourceFile)
sourceFile.close()
os.system("ssmtp " + para + " < email.txt")