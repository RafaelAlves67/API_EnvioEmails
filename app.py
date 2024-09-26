import pandas as pd
import pyodbc 
import datetime
import smtplib
import email.message
from dotenv import load_dotenv 
load_dotenv() 
import os 

# credenciais do banco 
db_username = os.getenv('USERNAME')
db_password = os.getenv('PASSWORD')
server = os.getenv('DB_HOST')
database = os.getenv('DB_NAME')


# conexão com banco 
db = pyodbc.connect(f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};UID={db_username};PWD={db_password}')
print("Banco de dados SQL SERVER CONECTADO.") 

# query que vai rodar
query = """ 
    SELECT 
    R.EmpCod, 
    R.CCtrlCodEstr, 
    CC.CCtrlNome, 
    R.ReqCompNum, 
    R.ReqCompDescr, 
    R.ReqCompData, 
    R.ReqCompStat, 
    R.ReqCompAprov,  
    R.FuncCod, 
    R.UsuCod, 
    GU.UsuCod AS GestorCod,
    U.UsuNome AS GestorNome,
    U.UsuEmail AS GestorEmail,
    U.UsuSenhaInternet  AS EmailSenha
FROM 
    REQ_COMP R 
INNER JOIN 
    CENTRO_CTRL CC ON CC.CCtrlCodEstr = R.CCtrlCodEstr
INNER JOIN 
    (SELECT DISTINCT 
         GU1.UsuCod, 
         GU1.GrpUsuCod
     FROM 
         GRP_X_USUARIO GU1
     WHERE 
         GU1.GrpUsuSuperv = 'T') AS GU 
    ON GU.GrpUsuCod = (
        SELECT TOP 1 
            GU2.GrpUsuCod 
        FROM 
            GRP_X_USUARIO GU2 
        WHERE 
            GU2.UsuCod = R.UsuCod 
            AND (GU2.GrpUsuCod LIKE '%req compras%' and GU2.GrpUsuCod LIKE '%labcor%')
    )
INNER JOIN 
    USUARIO U ON U.UsuCod = GU.UsuCod
WHERE 
    R.ReqCompAprov = 'Não'
    and R.ReqCompReprov = 'Não'
    AND R.EmpCod IN ('1.03.03', '1.03.01')
   AND R.ReqCompData > '2024-07-15'
   AND U.UsuEmail = 'rafael.alves@labcor.com.br'
ORDER BY 
    R.ReqCompData DESC;
"""

# pegando os resultados da query
result = pd.read_sql(query, db) 
result



def enviar_email(gestor_email, gestor_nome, reqcomp):
    # Configuração do servidor SMTP
    remetente = os.getenv('EMAIL')
    password = os.getenv('EMAIL_PASSWORD')

    # corpo do email 
    corpo_email = f"""
        Olá {gestor_nome},

        Existe uma nova requisição de compra com número ${reqcomp} aguardando sua aprovação. 

Link para aprovar: https://alvoerplabcor.zammi.com.br/#/alvoerp/reqcomp`
"""

    try:
        msg = email.message.Message() 
        msg['Subject'] = f"Requisição {reqcomp} aguardando aprovação"
        msg['From'] = remetente
        msg['To'] = gestor_email 
        msg.add_header('Content-Type', 'text/html')
        msg.set_payload(corpo_email)  

        s = smtplib.SMTP('smtp.office365.com: 587') 
        s.starttls() 

        # login
        s.login(msg['From'], password) 
        s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8')) 
        print(f'Email enviado para {gestor_email} para requisição {reqcomp}')

    except Exception as e: 
        print(f'Falha ao enviar e-email para {gestor_email}: {str(e)}')
    


# iterar sobre o resultado da query
for index, row in result.iterrows():
    enviar_email(
        gestor_email=row['GestorEmail'],
        reqcomp=row['ReqCompNum'],
        gestor_nome=row['GestorNome'],
    )