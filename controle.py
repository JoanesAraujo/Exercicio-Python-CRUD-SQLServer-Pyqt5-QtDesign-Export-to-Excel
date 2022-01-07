from PyQt5 import  uic,QtWidgets
from PyQt5.QtWidgets import QHBoxLayout, QMessageBox, QPushButton, QTableWidget, QVBoxLayout, QWidget
from PyQt5.QtGui import QIcon
from reportlab.pdfgen import canvas
import pyodbc
import pandas as pd
import os
from datetime import datetime

dados_conexao = (
    "Driver={SQL Server};"
    "Server=DESKTOP-5SO7F5F;"
    "Database=pessoa;"
)

conexao = pyodbc.connect(dados_conexao)
print("Conexão Bem Sucedida")

def refresh():
    cursor = conexao.cursor()
    cursor.execute("SELECT * FROM pes_fisica")
    dados_lidos = cursor.fetchall()
    print(dados_lidos)
    # Alinhamento da tabela de listar
    cadastro.tableWidget.setColumnWidth (0, 40 )
    cadastro.tableWidget.setColumnWidth (1, 100 )
    cadastro.tableWidget.setColumnWidth (4, 277 )
    # fim alinhamento
    cadastro.tableWidget.setRowCount(len(dados_lidos))
    cadastro.tableWidget.setColumnCount(5)

    for i in range(0, len(dados_lidos)):
            for j in range(0, 5):
                cadastro.tableWidget.setItem(i,j,QtWidgets.QTableWidgetItem(str(dados_lidos[i][j])))
    #-FIM LISTA-

def export_to_excel():
    cursor = conexao.cursor()
    sqlQuery = "SELECT * FROM dbo.pes_fisica"
    cursor.execute(sqlQuery)
    dados_lidos = cursor.fetchall()

    writer = pd.ExcelWriter("LISTA_DE_PESSOAS.xlsx", engine = 'xlsxwriter')
    # here, you can store your query's result data
    df = pd.read_sql(sql = sqlQuery, con = conexao)
    df.to_excel(writer, sheet_name = 'Pessoas_Fisicas')

    writer.save()
    #writer.close()
    
    
    # mensagem ao salvar com sucesso
    msg = QMessageBox()
    msg.setIcon(msg.Information)
    msg.setWindowTitle("Sucesso")
    msg.setText("Arquivo Exportado com Sucesso!" )
    msg.exec()
    # fim mensagem

def excluir():
    linha = cadastro.tableWidget.currentRow()
    #print(linha)
    cadastro.tableWidget.removeRow(linha)
    cursor = conexao.cursor()
    cursor.execute("SELECT id FROM pes_fisica")
    dados_lidos = cursor.fetchall()
    valor_id = dados_lidos[linha][0]
    cursor.execute("DELETE FROM pes_fisica WHERE id="+ str(valor_id))

    cursor.commit()
    
def funcao_principal():
    linha1 = cadastro.lineName.text()
    linha2 = cadastro.lineCPF.text()
    linha3 = cadastro.lineCel.text()
    linha4 = cadastro.lineEmail.text()
    
       
    cursor = conexao.cursor()
    comando = f"""INSERT INTO pes_fisica (nome, cpf, cel, email)
VALUES
    ('{linha1}', '{linha2}', '{linha3}', '{linha4}')"""

    cursor.execute(comando)
    cursor.commit()

    # mensagem ao salvar com sucesso
    msg = QMessageBox()
    msg.setIcon(msg.Information)
    msg.setWindowTitle("Sucesso")
    msg.setText(" Usuario salvo com sucesso" )
    msg.exec()
    # fim mensagem

    # limpar campos ao adicionar
    cadastro.lineName.setText("")
    cadastro.lineCPF.setText("")
    cadastro.lineCel.setText("")
    cadastro.lineEmail.setText("")
    # fim limpar campos
    

app=QtWidgets.QApplication([])
cadastro=uic.loadUi("cadastro.ui")
cadastro.setWindowTitle("CADASTRO PESSOA FÍSICA")
cadastro.pushButton.clicked.connect(funcao_principal)
cadastro.pushButton_2.clicked.connect(refresh)
cadastro.pushButton_3.clicked.connect(export_to_excel)
cadastro.pushButton_4.clicked.connect(excluir)





#-LISTA DE PESSOAS-
cursor = conexao.cursor()
cursor.execute("SELECT * FROM pes_fisica")
dados_lidos = cursor.fetchall()
# Alinhamento da tabela de listar
cadastro.tableWidget.setColumnWidth (0, 40 )
cadastro.tableWidget.setColumnWidth (1, 100 )
cadastro.tableWidget.setColumnWidth (4, 277 )
# fim alinhamento
cadastro.tableWidget.setRowCount(len(dados_lidos))
cadastro.tableWidget.setColumnCount(5)

for i in range(0, len(dados_lidos)):
        for j in range(0, 5):
           cadastro.tableWidget.setItem(i,j,QtWidgets.QTableWidgetItem(str(dados_lidos[i][j])))
#-FIM LISTA-

cadastro.show()
app.exec()

