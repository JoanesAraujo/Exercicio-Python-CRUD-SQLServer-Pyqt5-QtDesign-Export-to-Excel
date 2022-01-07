from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QApplication, QTableWidgetSelectionRange, QWidget,\
    QVBoxLayout, QTableWidgetItem, QTableWidget
import sys
import pyodbc 
from PyQt5.QtGui import QIcon

dados_conexao = (
    "Driver={SQL Server};"
    "Server=DESKTOP-5SO7F5F;"
    "Database=pessoa;"
)

conexao = pyodbc.connect(dados_conexao)
print("Conexão Bem Sucedida")

class Window(QWidget):
    def __init__(self):
        super().__init__()


        self.setGeometry(200,200,400,200)
        self.setWindowTitle("Creating TableWidget")
        self.setWindowIcon(QIcon("python.png"))

        self.create_tables()



    def create_tables(self):
        vbox = QVBoxLayout()

        cursor = conexao.cursor()
        cursor.execute("SELECT * FROM pes_fisica")
        dados_lidos = cursor.fetchall()
        print(dados_lidos)


        table_widget = QTableWidget()
        table_widget.setRowCount(5)
        table_widget.setColumnCount(5)

        for i in range(0, len(dados_lidos)):
            for j in range(0, 5):
                table_widget.setItem(i,j,QtWidgets.QTableWidgetItem(str(dados_lidos[i][j]))) 


        


        vbox.addWidget(table_widget)


        self.setLayout(vbox)



App = QApplication(sys.argv)
window = Window()
window.show()
sys.exit(App.exec())