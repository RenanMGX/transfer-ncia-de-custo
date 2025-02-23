# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file '.\untitled.ui'
#
# Created by: PyQt5 UI code generator 5.15.10
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets
from bot_tranferencia_custo import Robo, Config
from tkinter import filedialog
from datetime import datetime
import os
from getpass import getuser




class Ui_title(QtWidgets.QDialog):
    def __init__(self):
        super().__init__()
        
        self.versao = "V1.2"
        self.setObjectName("title")
        width, Height = 450, 450
        self.resize(width, Height)
        
        self.setMaximumWidth(width)
        self.setMinimumWidth(width)
        
        self.setMaximumHeight(Height)
        self.setMinimumHeight(Height)
        
        self.horizontalLayoutWidget = QtWidgets.QWidget(self)
        self.horizontalLayoutWidget.setGeometry(QtCore.QRect(10, 10, 721, 75))
        self.horizontalLayoutWidget.setObjectName("horizontalLayoutWidget")
        self.horizontalLayout = QtWidgets.QHBoxLayout(self.horizontalLayoutWidget)
        self.horizontalLayout.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout.setObjectName("horizontalLayout")
        #spacerItem = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        #self.horizontalLayout.addItem(spacerItem)


        self.update_base_button = QtWidgets.QPushButton(self.horizontalLayoutWidget)
        self.update_base_button.setObjectName("update_base_button")
        self.update_base_button.clicked.connect(self.atualizar_base)
        self.horizontalLayout.addWidget(self.update_base_button)

        self.base_header = QtWidgets.QListWidget(self.horizontalLayoutWidget)
        self.base_header.setObjectName("base_header")
        item = QtWidgets.QListWidgetItem()
        self.base_header.addItem(item)
        self.horizontalLayout.addWidget(self.base_header)
        spacerItem1 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout.addItem(spacerItem1)
        
        self.horizontalLayoutWidget_2 = QtWidgets.QWidget(self)
        self.horizontalLayoutWidget_2.setGeometry(QtCore.QRect(10, 90, 421, 550))
        self.horizontalLayoutWidget_2.setObjectName("horizontalLayoutWidget_2")
        
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout(self.horizontalLayoutWidget_2)
        self.horizontalLayout_2.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        spacerItem2 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_2.addItem(spacerItem2)
        
        self.calendar = QtWidgets.QCalendarWidget(self.horizontalLayoutWidget_2)
        self.calendar.setGeometry(0,10,420,200)
        self.calendar.setVisible(True)

        self.load_file_button = QtWidgets.QPushButton(self.horizontalLayoutWidget_2)
        self.load_file_button.setObjectName("load_file_button")
        self.load_file_button.setVisible(False)
        self.load_file_button.clicked.connect(self.inicar_bot)
        #self.load_file_button.clicked.connect(self.test)
        
        self.horizontalLayout_2.addWidget(self.load_file_button)

        self.text_area = QtWidgets.QTextEdit(self.horizontalLayoutWidget_2)
        self.text_area.setText("")
        self.text_area.setVisible(False)
        self.text_area.setReadOnly(True)
        self.text_area.setGeometry(15,300,400,50)
        
        
        #self.horizontalLayout_2.addWidget(self.text_area)

        spacerItem3 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_2.addItem(spacerItem3)

        self.retranslateUi()
        self.update_inter()
        QtCore.QMetaObject.connectSlotsByName(self)


    def retranslateUi(self):
        _translate = QtCore.QCoreApplication.translate
        self.setWindowTitle(_translate("title", f"{self.versao} | Bot Transferencia de Custos"))
        self.update_base_button.setText(_translate("title", "Atualizar Base"))
        __sortingEnabled = self.base_header.isSortingEnabled()
        self.base_header.setSortingEnabled(False)
        item = self.base_header.item(0)

        base = configura.load()['cadastro_de_empresas']  #type: ignore
        if base == "":
            base = "Base não Identificada!"
        item.setText(_translate("title", base))  #type: ignore

        self.base_header.setSortingEnabled(__sortingEnabled)
        self.load_file_button.setText(_translate("title", "Carregar Arquivo"))

    def atualizar_base(self):
        self.update_inter()
        caminho = filedialog.askopenfilename()
        if (".xlsx" in caminho.lower()) or (".xlsm" in caminho.lower()) or (".xlsb" in caminho.lower()) or (".xltx" in caminho.lower()):
            configura.update("cadastro_de_empresas",caminho)
            self.update_inter()
        else:
            print("não é excell")


    def update_inter(self):
        configura.check()
        _translate = QtCore.QCoreApplication.translate
        item = self.base_header.item(0)
        item.setText(_translate("title", configura.load()['cadastro_de_empresas']))  #type: ignore

        if configura.load()['cadastro_de_empresas'] != "":  #type: ignore
            self.load_file_button.setVisible(True)
        else:
            self.load_file_button.setVisible(False)
        

        if len(robo.arquivos_com_error) > 0:
            texto = ""
            for error,descri in robo.arquivos_com_error.items():
                #self.text_area.setText(f"{error} : {descri}\n")
                texto += f"{error} : {descri}\n"

            self.text_area.setText(texto)
            self.text_area.setVisible(True)
        else:
            self.text_area.setVisible(False)
            self.text_area.setText("")
        

    def inicar_bot(self):
        self.update_inter()
        
        data = self.calendar.selectedDate()
        robo.date = datetime(data.year(),data.month(), data.day())
        
        robo.carregar_cadastro_de_empresas()
    
        robo.listar_arquivos()
    
        robo.carregar_arquivos_da_lista()
        robo.salvar_planilha()
        
        print("Fim")
        self.update_inter()
        robo.arquivos_com_error = {}
        
    def test(self):
        print(ui.calendar.selectedDate())



if __name__ == "__main__":
    import sys
    import multiprocessing
    multiprocessing.freeze_support()
    try:
        #configuracoes = Config()
        configura = Config()
        robo = Robo(configura)

        app = QtWidgets.QApplication(sys.argv)
        ui = Ui_title()
        ui.show()
        
        try:
            sys.exit(app.exec_())
        except:
            pass
    except:
        import traceback
        erro = traceback.format_exc()
        print(erro)
        path = f"C:\\Users\\{getuser()}\\.bot_transferencia_custo\\logs\\"
        if not os.path.exists(path):
            os.makedirs(path)
            
        file_name = f"{path}log_error_{datetime.now().strftime('%d%m%Y%H%M%S')}.txt"
        with open(file_name, 'w', encoding='utf-8')as _file:
            _file.write(erro)
        
        input()
