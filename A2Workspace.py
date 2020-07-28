#Importing required libraries and modules
from flask import Flask, render_template, flash, request,redirect,url_for,send_file,send_from_directory,abort
import pyodbc 
import pandas as pd
import win32com.client
from pywintypes import com_error
#import win32com.client
#from pywintypes import com_error
from xlsxwriter.utility import xl_rowcol_to_cell
from datetime import datetime
import sys
import time

#################################################################################################

#Before deploying service:

#Step1: Install virtual environment
#Windows:
# py -m pip install --user virtualenv
#Linux/MacOS:
#python3 -m pip install --user virtualenv

#Step 2: Create an environment
#Windows:
# Navigate to service folder and execute: py -m venv env
#Linux:
# python3 -m venv env

#Step3: Activate the environment
#Windows:
# .\env\Scripts\activate
#MacOS/Linux:
#source env/bin/activate

#Step 4: Install all required python stuff

#must run thru terminal before deploying code:
#pip3 install Flask
#pip3 install pandas
#pip3 install pyodbc
#pip3 install openpyxl
#pip3 install pywin32
#pip3 install xlsxwriter

#Step5: Launch the app:
# py A2Workspace.py

#################################################

#Before running the script:

#navigate to .py directory
#run venv\Scripts\activate

#################################################

app = Flask(__name__)
log = 0

@app.route('/', methods=['POST','GET'])
def index():
    if log == 1:
        return render_template('index.html')
    else:
        return redirect('/login')
    #this will be the home and where you can select all options

# Route for handling the login page logic
@app.route('/login', methods=['POST','GET'])
def login():
    global log
    error = None
    if request.method == 'POST':
        if request.form['user'] != 'admin' or request.form['password'] != 'admin':
            error = 'Credenciales invalidas.'
        else:
            log = 1
            return redirect('/')
    return render_template('login.html', error=error)

@app.route("/modify_document", methods=["GET", "POST"])
def modify_document():
    operation = 0
    if log==1:
        #any POST submitted? Capture params and perform task
        if request.method == 'POST':
            #capture parameters sent from from
            num1 = request.form['num1']
            num2 = request.form['num2']
            num3 = request.form['num3']
            if (len(num1)==8 and len(num2)==8 and len(num3)==5):
                connect,cursor = connection(1)
                connect,cursor = connection(1)
                query1 = query_construction(ccode1=num2,ccode2=num1,ccode3=num3,query_type=1)
                operation = querying_dbisam(query1=query1,connect=connect,cursor=cursor,query_type=1)
                if operation == 1:
                    return redirect('/query_exception')
                if operation == 0:
                    return redirect('/query_ok')
            else:
                return redirect('/query_exception')
        #If no form has been submitted then render form
        else:
            return render_template("modify_document.html")
    else:
        return redirect('/login')

@app.route("/modify_date", methods=["GET", "POST"])
def modify_date():
    #any POST submitted? Capture params and perform task
    if log==1:
        if request.method == 'POST':
            #capture parameters sent from from
            num1 = request.form['num1']
            date1 = request.form['date1']
            if len(num1)==8:
                connect,cursor = connection(1)
                query1,query2 = query_construction(ccode1=num1,ccode2=date1,ccode3=None,query_type=4)
                operation = querying_dbisam(query1=query1,connect=connect,cursor=cursor,query_type=1)
                if operation == 1:
                    return redirect('/query_exception')
                if operation == 0:
                    return redirect('/query_ok')
            else:
                return redirect('/query_exception')
        else:
            #If no form has been submitted then render form
            return render_template("modify_date.html")
    else:
        return redirect('login')

@app.route("/query_document", methods=["GET", "POST"])
def query_document():
    if log==1:
        #any POST submitted? Capture params and perform task
        if request.method == 'POST':
            #capture parameters sent from from
            num1 = request.form['num1']
            if len(num1)==5:
                try:
                    connect = connection(0)
                    query1,query2 = query_construction(ccode1=num1,ccode2='0',ccode3=None,query_type=3)
                    result = querying_dbisam(query1=query1,connect=connect,query_type=3)
                    result = result.rename(columns={'FCC_CODIGO':'# Afiliado','FCC_NUMERO':'# Documento','FCC_FECHAEMISION':'Fecha Emision','FCC_TIPOTRANSACCION':'Tipo Transaccion','FCC_DESCRIPCIONMOV':'Descripcion','FCC_MONTODOCUMENTO':'Monto','FCC_SALDODOCUMENTO':'Saldo Documento'})
                    result = result.replace('\n','',regex=True).replace('\r','',regex=True)
                    if len(result)>0:
                        return render_template("document.html",tables=[result.to_html(classes='data')],titles=result.columns.values)
                    else:
                        return redirect('/invalid_query')
                except Exception as e:
                    return redirect('/invalid_query')
            else:
                return redirect('/invalid_query')
        else:
            #If no form has been submitted then render form
            return render_template("query_document.html")
    else:
        return redirect('login')

@app.route("/query_affiliate", methods=["GET", "POST"])
def query_affiliate():
    if log==1:
        #any POST submitted? Capture params and perform task
        if request.method == 'POST':
            #capture parameters sent from from
            num1 = request.form['num1']
            if len(num1)==5:
                try:
                    connect = connection(0)
                    query1,query2 = query_construction(ccode1='0',ccode2=num1,ccode3=None,query_type=2)
                    result = querying_dbisam(query2=query2,connect=connect,query_type=2)
                    result = result.rename(columns={'FC_CODIGO':'# Cliente','FC_DESCRIPCION':'Nombre','FC_STATUS':'Estado','FC_DIRECCION1':'Direccion','FC_TELEFONO':'Telefono','FC_EMAIL':'Email'})
                    result['Estado'] = result['Estado'].apply(lambda x: 'Activo' if x == True else 'Inactivo')
                    if len(result)>0:
                        return render_template("affiliate.html",tables=[result.to_html(classes='data')],titles=result.columns.values)
                    else:
                        return redirect('/invalid_query')
                except Exception as e:
                    return redirect('/invalid_query')
            else:
                return redirect('/invalid_query')
        else:
            #If no form has been submitted then render form
            return render_template("query_affiliate.html")
    else:
        return redirect('/login')


@app.route("/get_account_state", methods=["GET", "POST"])
def send_account_state():
    #any GET submitted? Capture params and perform task
    if request.method == 'GET':
        #capture parameters sent from form
        try:
            num1 = request.args.get('num1', None)
        except Exception as e:
            abort(404)
        if len(num1)==5:
            try:
                connect = connection(0)
                query1,query2 = query_construction(ccode1=num1,ccode2=None,query_type=0)
                acc_query,client_query = querying_dbisam(query1=query1,query2=query2,connect=connect,query_type=0)
                if len(client_query)>0:
                    pass
                else:
                    abort(404)
                #processing the previous datasets in order to remove what's not needed and also to get the account balance result
                acc_balance_df,acc_result,client_result = data_processing(acc_query,client_query)
                #creating an excel file with the resulting dataframes
                build_excel(acc_balance_df,acc_result,client_result,num1)
                #transforming excel file to pdf format so that's the business requirement.
                excel_to_pdf(num1)
                filepath = r'C:/Users/MARIA SANCHEZ/Documents/PythonProjects/A2Workspace/temp/'+num1+'.pdf'
                return send_file(filepath)
            except Exception as e:
                abort(404)
        else:
            abort(404)

    #If no form has been submitted then crash
    abort(404)


@app.route("/query_exception")
def query_exception():
    return render_template("query_exception.html")

@app.route("/query_ok")
def query_ok():
    return render_template("query_ok.html")

@app.route("/invalid_query")
def invalid_query():
    return render_template("invalid_query.html")

@app.route("/exit")
def exit():
    global log
    log=0
    return redirect("/login")

def connection(connection_type):

    """This function handles the pyodbc.connect and returns a 'connect' object used to query DBISAM"""
    if connection_type==0:
        #Connection to DBISAM through ODBC to only read
        connect = pyodbc.connect("DRIVER={DBISAM 4 ODBC Driver};ConnectionType=Local;CatalogName=C:/a2AdminCCN/Empre001/Data;")
        return connect
    if connection_type==1:
        #Connection to DBISAM through ODBC to r/w
        connect = pyodbc.connect("DRIVER={DBISAM 4 ODBC Driver};ConnectionType=Local;CatalogName=C:/a2AdminCCN/Empre001/Data;ReadOnly=False")
        cursor = connect.cursor()
        return connect,cursor

def query_construction(ccode1=None,ccode2=None,ccode3=None,query_type=None):

    """This function builds the queries to then extract the required data from DBISAM"""
    if query_type==0:
        #Queries' construction
        code = "'"+ccode1+"'"
        query1 = 'SELECT FCC_CODIGO,FCC_NUMERO,FCC_FECHAEMISION,FCC_TIPOTRANSACCION,FCC_DESCRIPCIONMOV,FCC_MONTODOCUMENTO,FCC_SALDODOCUMENTO FROM Scuentasxcobrar WHERE FCC_CODIGO LIKE '
        query1 = query1+code
        query2 = 'SELECT FC_CODIGO,FC_DESCRIPCION,FC_STATUS,FC_DIRECCION1,FC_TELEFONO,FC_EMAIL FROM Sclientes WHERE FC_CODIGO LIKE '
        query2 = query2+code
        return query1,query2
    if query_type==1:
        code1 = "'"+ccode1+"'"
        code2 = "'"+ccode2+"'"
        code3 = "'"+ccode3+"'"
        #query1 = 'SELECT FCC_CODIGO, FCC_TIPOTRANSACCION FROM Scuentasxcobrar WHERE FCC_TIPOTRANSACCION=6 AND FCC_NUMERO='+code1+' UPDATE Scuentasxcobrar SET FCC_NUMERO='+code2+' WHERE FCC_NUMERO='+code1
        query1 = 'UPDATE Scuentasxcobrar SET FCC_NUMERO='+code1+' WHERE FCC_NUMERO='+code2+' AND FCC_CODIGO='+code3+' AND FCC_TIPOTRANSACCION=6'
        return query1
    if query_type==2:
        #Queries' construction
        code = "'"+ccode2+"'"
        query1 = 'SELECT FCC_CODIGO,FCC_NUMERO,FCC_FECHAEMISION,FCC_TIPOTRANSACCION,FCC_DESCRIPCIONMOV,FCC_MONTODOCUMENTO,FCC_SALDODOCUMENTO FROM Scuentasxcobrar WHERE FCC_CODIGO LIKE '
        query1 = query1+code
        query2 = 'SELECT FC_CODIGO,FC_DESCRIPCION,FC_STATUS,FC_DIRECCION1,FC_TELEFONO,FC_EMAIL FROM Sclientes WHERE FC_CODIGO LIKE '
        query2 = query2+code
        return query1,query2
    if query_type==3:
        #Queries' construction
        code = "'"+ccode1+"'"
        query1 = 'SELECT FCC_CODIGO,FCC_NUMERO,FCC_FECHAEMISION,FCC_TIPOTRANSACCION,FCC_DESCRIPCIONMOV,FCC_MONTODOCUMENTO,FCC_SALDODOCUMENTO FROM Scuentasxcobrar WHERE FCC_CODIGO LIKE '
        query1 = query1+code
        query2 = None
        return query1,query2
    if query_type==4:
        #Queries' construction
        code = "'"+ccode1+"'"
        date = "'"+ccode2+"'"
        query1 = 'UPDATE Scuentasxcobrar SET FCC_FECHAEMISION='+date+' WHERE FCC_NUMERO='+code
        query2 = None
        return query1,query2

def querying_dbisam(query1=None,query2=None,connect=None,cursor=None,query_type=None):

    """This function queries DBISAM using previous strings built at query_construction function."""
    if query_type==0:
        #Querying 'SCuentasxcobrar' and 'SClientes' data with Pandas
        acc_query = pd.read_sql_query(query1, connect)
        client_query = pd.read_sql_query(query2, connect)
        return acc_query,client_query
    if query_type==1:
        #Querying 'SCuentasxcobrar'
        try:
            cursor.execute(query1)
            connect.commit()
        except Exception as e:
            return 1
        return 0
    if query_type==2:
        #Querying 'SClientes' data with Pandas
        client_query = pd.read_sql_query(query2, connect)
        return client_query
    if query_type==3:
        #Querying 'SCuentasxcobrar' data with Pandas
        acc_query = pd.read_sql_query(query1, connect)
        return acc_query

def data_processing(acc_query,client_query):

    """This function takes query results and packs it into Pandas dataframes to process the data and later on build the files"""
    #Creating DataFrames with the previous query results
    acc_result = pd.DataFrame(acc_query)
    client_result = pd.DataFrame(client_query)
    #After the query, FCC_CODIGO is not useful, hence, dropping.
    del acc_result['FCC_CODIGO']
    #Renaming columns so they are more friendly to users
    client_result = client_result.rename(columns={'FC_CODIGO':'# Cliente','FC_DESCRIPCION':'Nombre','FC_STATUS':'Estado','FC_DIRECCION1':'Direccion','FC_TELEFONO':'Telefono','FC_EMAIL':'Email'})
    acc_result = acc_result.rename(columns={'FCC_NUMERO':'# Documento','FCC_FECHAEMISION':'Fecha Emision','FCC_TIPOTRANSACCION':'Tipo Transaccion','FCC_DESCRIPCIONMOV':'Descripcion','FCC_MONTODOCUMENTO':'Monto','FCC_SALDODOCUMENTO':'Saldo Documento'})
    #Renaming 'Estado' values to more appropiate ones
    client_result['Estado'] = client_result['Estado'].apply(lambda x: 'Activo' if x == True else 'Inactivo')
    #Renaming 'Tipo Transaccion' values to more meaningful ones
    acc_result['Tipo Transaccion'] = acc_result['Tipo Transaccion'].apply(lambda x: 'Factura' if x == 1 else x)
    acc_result['Tipo Transaccion'] = acc_result['Tipo Transaccion'].apply(lambda x: 'Nota debito' if x == 2 else x)
    acc_result['Tipo Transaccion'] = acc_result['Tipo Transaccion'].apply(lambda x: 'Pago' if x == 4 else x)
    acc_result['Tipo Transaccion'] = acc_result['Tipo Transaccion'].apply(lambda x: 'Nota credito' if x == 5 else x)
    acc_result['Tipo Transaccion'] = acc_result['Tipo Transaccion'].apply(lambda x: 'Adelanto' if x == 6 else x)
    acc_result['Tipo Transaccion'] = acc_result['Tipo Transaccion'].apply(lambda x: 'NC por adelanto' if x == 9 else x)
    #Dropping 'Tipo Transaccion' == 54 as per it won't add any value to the final result
    acc_result = acc_result[acc_result['Tipo Transaccion']!=54]
    #Setting proper indexes
    client_result = client_result.set_index('# Cliente')
    acc_result = acc_result.set_index('# Documento')
    #Calculating account balance
    acc_balance = float(0)
    for x in range(len(acc_result)):
        if acc_result.iloc[x]['Tipo Transaccion'] == 'Factura':
            acc_balance -= acc_result.iloc[x]['Monto']
        if acc_result.iloc[x]['Tipo Transaccion'] == 'Nota debito':
            acc_balance -= acc_result.iloc[x]['Monto']
        if acc_result.iloc[x]['Tipo Transaccion'] == 'Pago':
            acc_balance += acc_result.iloc[x]['Monto']
        if acc_result.iloc[x]['Tipo Transaccion'] == 'Nota credito':
            acc_balance += acc_result.iloc[x]['Monto']
        if acc_result.iloc[x]['Tipo Transaccion'] == 'Adelanto':
            acc_balance += acc_result.iloc[x]['Monto']    
    #Building account balance DataFrame to then insert it into the excel file
    acc_balance_df = pd.Series({'Saldo cuenta':acc_balance})
    acc_balance_df = pd.DataFrame(acc_balance_df).transpose()
    #Setting DF style
    acc_balance_df = acc_balance_df.style.set_properties(**{'background-color': 'white','font-size': '10pt'})

    return acc_balance_df,acc_result,client_result

def build_excel(acc_balance_df,acc_result,client_result,client_number):

    """This function takes three dataframes resulting from the previous processing and buils an excel file with them."""
    #Saving DataFrames to an Excel File
    filepath = 'temp/'+client_number+'.xlsx'
    with pd.ExcelWriter(filepath, engine='xlsxwriter') as writer:
        client_result.to_excel(writer, sheet_name='Sheet1', startcol=1,startrow=6)
        acc_balance_df.to_excel(writer, sheet_name='Sheet1',startcol=1, startrow=9)
        acc_result.to_excel(writer, sheet_name='Sheet1',startcol=1, startrow=12)
        #Adding style and format
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        font_fmt = workbook.add_format({'font_name': 'Calibri', 'font_size': 10,'bold': False})
        font_fmt1 = workbook.add_format({'font_name': 'Calibri', 'font_size': 12,'bold': False})
        font_fmt2 = workbook.add_format({'font_name': 'Calibri', 'font_size': 13,'bold': True})
        font_fmt3 = workbook.add_format({'font_name': 'Calibri', 'font_size': 12,'bold': False,'align':'right'})
        worksheet.set_landscape()
        worksheet.set_column('A:A', 0,font_fmt)
        worksheet.set_column('B:B', 15,font_fmt)
        worksheet.set_column('C:C', 20,font_fmt)
        worksheet.set_column('D:D', 15,font_fmt)
        worksheet.set_column('E:E', 30,font_fmt)
        worksheet.set_column('F:F', 15,font_fmt)
        worksheet.set_column('G:G', 20,font_fmt)
        #Adding header and giving its format
        worksheet.write('B2', 'Colegio de Contadores Públicos de Nicaragua',font_fmt2)
        worksheet.write('B3', 'Estado de cuenta',font_fmt1)
        #worksheet.write('B3', 'Cuentas por cobrar',font_fmt1)
        worksheet.write('G2', datetime.today().strftime('%d-%m-%Y'),font_fmt3)
        writer.save()

        return 

def excel_to_pdf(client_number):

    """This function takes the excel file and transforms it into a pdf one which will be the final product."""
    # Path to original excel file
    WB_PATH = r'C:/Users/MARIA SANCHEZ/Documents/PythonProjects/A2Workspace/temp/'+client_number+'.xlsx'
    # PDF path when saving
    PATH_TO_PDF = 'PythonProjects/A2Workspace/temp/'+client_number+'.pdf'
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    try:
        print('Start conversion to PDF')
        # Open excel file
        wb = excel.Workbooks.Open(WB_PATH)
        # Specifying the sheet that will be converted to PDF format
        ws_index_list = [1]
        wb.WorkSheets(ws_index_list).Select()
        # Save
        wb.ActiveSheet.ExportAsFixedFormat(0, PATH_TO_PDF)
    except com_error as e:
        print('failed.')
        abort(404)
    else:
        print('Succeeded.')
    finally:
        wb.Close(True)
        excel.Quit()

    return

if __name__ == '__main__':
    app.run(debug=True)