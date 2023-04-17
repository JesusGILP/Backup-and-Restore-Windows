############################################################
#
#  Software.py
#  Version 1.0
#  Created by: Jesús Gil Puerto
#  Date: 17/04/2023
#  Description: Software de copia de archivos
#
############################################################


import datetime
import os
import socket
import subprocess
import getpass
import time
from PyQt5.QtWidgets import QMessageBox
import zipfile
import win32print
import pywintypes
from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_SoftwareCopy(object):
    def setupUi(self, SoftwareCopy):
        SoftwareCopy.setObjectName("SoftwareCopy")
        SoftwareCopy.resize(787, 565)
        SoftwareCopy.setDocumentMode(False)
        self.centralwidget = QtWidgets.QWidget(SoftwareCopy)
        self.centralwidget.setObjectName("centralwidget")
        self.tabWidget = QtWidgets.QTabWidget(self.centralwidget)
        self.tabWidget.setGeometry(QtCore.QRect(20, 20, 741, 491))
        self.tabWidget.setObjectName("tabWidget")
        self.tab = QtWidgets.QWidget()
        self.tab.setObjectName("tab")
        self.groupBox = QtWidgets.QGroupBox(self.tab)
        self.groupBox.setGeometry(QtCore.QRect(10, 20, 711, 421))
        self.groupBox.setObjectName("groupBox")
        self.boton_copiar = QtWidgets.QPushButton(self.groupBox)
        self.boton_copiar.setGeometry(QtCore.QRect(600, 370, 101, 41))
        self.boton_copiar.setStyleSheet("background-color: rgb(53, 132, 228);")
        self.boton_copiar.setObjectName("boton_copiar")
        self.frame_copia = QtWidgets.QFrame(self.groupBox)
        self.frame_copia.setGeometry(QtCore.QRect(10, 70, 331, 291))
        self.frame_copia.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_copia.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_copia.setObjectName("frame_copia")
        #cambiamos el color del frame a gris claro background-color: rgb(246, 245, 244);
        self.frame_copia.setStyleSheet("background-color: rgb(246, 245, 244);")
        self.toolButton_copia = QtWidgets.QToolButton(self.groupBox)
        self.toolButton_copia.setGeometry(QtCore.QRect(10, 370, 331, 41))
        self.toolButton_copia.setStyleSheet("background-color: rgb(51, 209, 122);")
        self.toolButton_copia.setObjectName("toolButton_copia")
        self.label = QtWidgets.QLabel(self.groupBox)
        self.label.setGeometry(QtCore.QRect(370, 40, 321, 20))
        self.label.setObjectName("label")
        self.frame_copia_2 = QtWidgets.QFrame(self.groupBox)
        self.frame_copia_2.setGeometry(QtCore.QRect(370, 70, 331, 291))
        self.frame_copia_2.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_copia_2.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_copia_2.setObjectName("frame_copia_2")
        self.frame_copia_2.setStyleSheet("background-color: rgb(246, 245, 244);")
        self.label_3 = QtWidgets.QLabel(self.groupBox)
        self.label_3.setGeometry(QtCore.QRect(12, 40, 271, 20))
        self.label_3.setObjectName("label_3")
        self.tabWidget.addTab(self.tab, "")
        self.tab_2 = QtWidgets.QWidget()
        self.tab_2.setObjectName("tab_2")
        self.groupBox_3 = QtWidgets.QGroupBox(self.tab_2)
        self.groupBox_3.setGeometry(QtCore.QRect(10, 20, 711, 421))
        self.groupBox_3.setObjectName("groupBox_3")
        self.boton_restaurar = QtWidgets.QPushButton(self.groupBox_3)
        self.boton_restaurar.setGeometry(QtCore.QRect(600, 370, 101, 41))
        self.boton_restaurar.setStyleSheet("background-color: rgb(53, 132, 228);")
        self.boton_restaurar.setObjectName("boton_restaurar")
        self.frame_restaurar = QtWidgets.QFrame(self.groupBox_3)
        self.frame_restaurar.setGeometry(QtCore.QRect(10, 70, 331, 291))
        self.frame_restaurar.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_restaurar.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_restaurar.setObjectName("frame_restaurar")
        self.frame_restaurar.setStyleSheet("background-color: rgb(246, 245, 244);")
        self.toolButton_restaurar = QtWidgets.QToolButton(self.groupBox_3)
        self.toolButton_restaurar.setGeometry(QtCore.QRect(10, 370, 331, 41))
        self.toolButton_restaurar.setStyleSheet("background-color: rgb(51, 209, 122);")
        self.toolButton_restaurar.setObjectName("toolButton_restaurar")
        self.label_2 = QtWidgets.QLabel(self.groupBox_3)
        self.label_2.setGeometry(QtCore.QRect(370, 40, 321, 20))
        self.label_2.setObjectName("label_2")
        self.frame_restaurar_2 = QtWidgets.QFrame(self.groupBox_3)
        self.frame_restaurar_2.setGeometry(QtCore.QRect(370, 70, 331, 291))
        self.frame_restaurar_2.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_restaurar_2.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_restaurar_2.setObjectName("frame_restaurar_2")
        self.frame_restaurar_2.setStyleSheet("background-color: rgb(246, 245, 244);")
        self.label_5 = QtWidgets.QLabel(self.groupBox_3)
        self.label_5.setGeometry(QtCore.QRect(12, 40, 271, 20))
        self.label_5.setObjectName("label_5")
        self.tabWidget.addTab(self.tab_2, "")
        SoftwareCopy.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(SoftwareCopy)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 787, 22))
        self.menubar.setObjectName("menubar")
        SoftwareCopy.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(SoftwareCopy)
        self.statusbar.setObjectName("statusbar")
        SoftwareCopy.setStatusBar(self.statusbar)
        self.menubar = QtWidgets.QMenuBar(SoftwareCopy)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 787, 22))
        self.menubar.setObjectName("menubar")
        self.menumen = QtWidgets.QMenu(self.menubar)
        self.menumen.setObjectName("menumen")
        SoftwareCopy.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(SoftwareCopy)
        self.statusbar.setObjectName("statusbar")
        SoftwareCopy.setStatusBar(self.statusbar)
        self.actioninformaci_n = QtWidgets.QAction(SoftwareCopy)
        self.actioninformaci_n.setObjectName("actioninformaci_n")
        self.actionSalir = QtWidgets.QAction(SoftwareCopy)
        self.actionSalir.setObjectName("actionSalir")
        self.menumen.addSeparator()
        self.menumen.addAction(self.actioninformaci_n)
        self.menubar.addAction(self.menumen.menuAction())
        self.menumen.addSeparator()
        self.menumen.addAction(self.actionSalir)
        self.menubar.addAction(self.menumen.menuAction())



        self.retranslateUi(SoftwareCopy)
        self.tabWidget.setCurrentIndex(0)
        QtCore.QMetaObject.connectSlotsByName(SoftwareCopy)

    def retranslateUi(self, SoftwareCopy):
        _translate = QtCore.QCoreApplication.translate
        SoftwareCopy.setWindowTitle(_translate("SoftwareCopy", "Backup and Restore JS"))
        self.groupBox.setTitle(_translate("SoftwareCopy", "Copia de perfil"))
        self.boton_copiar.setText(_translate("SoftwareCopy", "Copiar"))
        self.toolButton_copia.setText(_translate("SoftwareCopy", "SELECCONE LA CARPETA DE COPIA"))
        self.label.setText(_translate("SoftwareCopy", "Carpeta de copia"))
        self.label_3.setText(_translate("SoftwareCopy", "Selecciona el perfil a copiar"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab), _translate("SoftwareCopy","Copia"))
        self.groupBox_3.setTitle(_translate("SoftwareCopy", "Restaurar perfil"))
        self.boton_restaurar.setText(_translate("SoftwareCopy", "Restaurar"))
        self.toolButton_restaurar.setText(_translate("SoftwareCopy", "SELECCONE LA CARPETA DE RESTAURACIÓN"))
        self.label_2.setText(_translate("SoftwareCopy", "Carpeta de restauración"))
        self.label_5.setText(_translate("SoftwareCopy", "Selecciona el perfil a restaurar"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab_2), _translate("SoftwareCopy", "Restaurar"))
        self.menumen.setTitle(_translate("SoftwareCopy", "menú"))
        self.actioninformaci_n.setText(_translate("SoftwareCopy", "información"))
        self.actionSalir.setText(_translate("SoftwareCopy", "Salir"))


class SoftwareCopyApp(QtWidgets.QMainWindow, Ui_SoftwareCopy):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setupUi(self)
        self.checkbox_perfiles_copia = []
        self.checkbox_perfiles_restaurar = []
        self.checkbox_carpetas_copia = []
        self.checkbox_zips_restaurar = []
        #la ruta de la carpeta de copia es el nombre del perfil seleccionado perfil_seleccionado
        self.carpetas = []
        # la ruta de la carpeta de copia de SAP ES AppData\Roaming\SAP\Common\
        self.carpetas.append("AppData\Roaming\SAP\Common")
        self.carpetas.append("AppData\Roaming\IBM\Personal Communications\private")
        self.carpetas.append("Documents")
        self.carpetas.append("Downloads")
        self.carpetas.append("Pictures")
        self.carpetas.append("Videos")
        self.carpetas.append("Desktop")
        self.carpetas.append("Favorites")
        self.carpetas.append("Music")


        self.crear_checkbox_perfiles()
        self.crear_checkbox_carpetas_copia()
        self.boton_copiar.clicked.connect(self.copiar_perfil)
        self.boton_restaurar.clicked.connect(self.restaurar_perfil)
        self.toolButton_copia.clicked.connect(self.seleccionar_carpeta_copia)
        self.toolButton_restaurar.clicked.connect(self.seleccionar_carpeta_restaurar)
        
        self.actionSalir.triggered.connect(self.salir)
        self.actioninformaci_n.triggered.connect(self.informacion)
        
    def informacion(self):
        #mostramos un pop up con la información del programa
        msg = QMessageBox()
        #tamano de la ventana
        msg.resize(400, 200)
        msg.setWindowTitle("Información del programa")
        #crear 2 saltos de linea
        msg.setText("Copia y Restauración de Perfiles Windows \n \n Desarrollado por: Jesús Gil Puerto \n \n Versión: 1.0")
        msg.setIcon(QMessageBox.Information)
        msg.setStandardButtons(QMessageBox.Ok)
        msg.exec_()

        
    def salir(self):
        self.close()
        
    def create_log(self, log_file_path):
        network_info = self.get_network_info()
        

        with open(log_file_path, 'w') as log_file:
            log_file.write(f'Machine Name: {network_info[0]}\n')
            log_file.write(f'User: {network_info[1]}\n')
            log_file.write(f'Network Configuration:\n{network_info[2]}\n')
        
    def get_network_info(self):
        machine_name = socket.gethostname()
        username = getpass.getuser()
        ipconfig_output = subprocess.check_output('ipconfig /all', shell=True, text=True)

        return machine_name, username, ipconfig_output
        
    def get_printer_ip(self,printer_name):
        try:
            hPrinter = win32print.OpenPrinter(printer_name)
            printer_info = win32print.GetPrinter(hPrinter, 2)
            win32print.ClosePrinter(hPrinter)
            port_name = printer_info['pPortName']
            
            if port_name.startswith('IP_'):
                return port_name[3:]
            else:
                return None
        except pywintypes.error:
            return None

    
    def save_printer_log(self, backup_imp):
        printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL)
        log_file = os.path.join(backup_imp)

        with open(log_file, "w") as f:
                for printer in printers:
                    name = printer[2]
                    ip = self.get_printer_ip(name)
                    if ip:
                        f.write(f"Nombre de la impresora: {name}, IP: {ip}\n")
                    else:
                        f.write(f"Nombre de la impresora: {name}, IP no disponible\n")


        return log_file
    
    def crear_checkbox_perfiles(self):
        perfiles = self.obtener_perfiles()
        for perfil in perfiles:
            cb_copia = QtWidgets.QCheckBox(self.frame_copia)
            cb_copia.setText(perfil)
            self.checkbox_perfiles_copia.append(cb_copia)
            cb_restaurar = QtWidgets.QCheckBox(self.frame_restaurar)
            cb_restaurar.setText(perfil)
            self.checkbox_perfiles_restaurar.append(cb_restaurar)

        # Acomodar los checkbox en el layout
        layout_copia = QtWidgets.QVBoxLayout(self.frame_copia)
        layout_restaurar = QtWidgets.QVBoxLayout(self.frame_restaurar)
        for cb_copia, cb_restaurar in zip(self.checkbox_perfiles_copia, self.checkbox_perfiles_restaurar):
            layout_copia.addWidget(cb_copia)
            layout_restaurar.addWidget(cb_restaurar)

    def obtener_perfiles(self):
        ruta_usuarios = os.path.expanduser("~").rsplit(os.sep, 1)[0]
        perfiles = []
        for elemento in os.listdir(ruta_usuarios):
            if os.path.isdir(os.path.join(ruta_usuarios, elemento)):
                perfiles.append(elemento)
        return perfiles

    def seleccionar_carpeta_copia(self):
        carpeta_copia = QtWidgets.QFileDialog.getExistingDirectory(None, "Seleccionar Carpeta", os.path.expanduser("~"))
        self.label_3.setText(carpeta_copia)
        
    def crear_checkbox_carpetas_copia(self):
        layout_copia_2 = QtWidgets.QVBoxLayout(self.frame_copia_2)
        for carpeta in self.carpetas:
            cb_copia = QtWidgets.QCheckBox(self.frame_copia_2)
            cb_copia.setText(carpeta)
            self.checkbox_carpetas_copia.append(cb_copia)
            layout_copia_2.addWidget(cb_copia)

    def actualizar_checkbox_zips(self):
        carpeta_restaurar = self.label_5.text()
        if os.path.exists(carpeta_restaurar):
            for cb in self.checkbox_zips_restaurar:
                cb.deleteLater()
            self.checkbox_zips_restaurar.clear()

            zips = [f for f in os.listdir(carpeta_restaurar) if f.endswith('.zip')]
            layout_restaurar_2 = QtWidgets.QVBoxLayout(self.frame_restaurar_2)
            for zip_file in zips:
                cb_zip = QtWidgets.QCheckBox(self.frame_restaurar_2)
                cb_zip.setText(zip_file)
                self.checkbox_zips_restaurar.append(cb_zip)
                layout_restaurar_2.addWidget(cb_zip)

    def seleccionar_carpeta_restaurar(self):
        carpeta_restaurar = QtWidgets.QFileDialog.getExistingDirectory(None, "Seleccionar Carpeta", os.path.expanduser("~"))
        self.label_5.setText(carpeta_restaurar)
        self.actualizar_checkbox_zips()

    def copiar_perfil(self):
        #obtener fecha y hora
        fecha = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        perfil_seleccionado = None
        for cb in self.checkbox_perfiles_copia:
            if cb.isChecked():
                perfil_seleccionado = cb.text()
                break

        if perfil_seleccionado is None:
            QtWidgets.QMessageBox.warning(self, "Error", "Por favor, seleccione un perfil para copiar.")
            return

        carpeta_copia = self.label_3.text()
        if not os.path.exists(carpeta_copia):
            QtWidgets.QMessageBox.warning(self, "Error", "Por favor, seleccione una carpeta de copia válida.")
            return

        ruta_usuarios = os.path.expanduser("~").rsplit(os.sep, 1)[0]
        ruta_perfil = os.path.join(ruta_usuarios, perfil_seleccionado)
        carpetas_seleccionadas = [cb.text() for cb in self.checkbox_carpetas_copia if cb.isChecked()]

        if not carpetas_seleccionadas:
            QtWidgets.QMessageBox.warning(self, "Error", "Por favor, seleccione al menos una carpeta para copiar.")
            return
        #tambien añadimos la fecha y hora al nombre del zip
        zip_path = os.path.join(carpeta_copia, f"{perfil_seleccionado}_{fecha}_Backup.zip")
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for carpeta in carpetas_seleccionadas:
                carpeta_path = os.path.join(ruta_perfil, carpeta)
                if os.path.exists(carpeta_path):
                            #no copiamos un archivo que se llama desktop.ini
                            #no copiamos un archivo que se llama thumbs.db
                            #no copiamos un archivo que se llama ntuser.dat
                            #no copiamos un archivo que se llama ntuser.ini
                        
                    for root, dirs, files in os.walk(carpeta_path):
                        for file in files:
                            if file not in ["desktop.ini", "thumbs.db", "ntuser.dat", "ntuser.ini"]:
                                file_path = os.path.join(root, file)
                                arcname = os.path.relpath(file_path, ruta_perfil)
                                zf.write(file_path, arcname)  

                else:
                    QtWidgets.QMessageBox.warning(self, "Error", f"No se encontró la carpeta {carpeta} en el perfil seleccionado.")

        QtWidgets.QMessageBox.information(self, "Copia realizada", "La copia de seguridad se realizó con éxito.")
        
        backup_folder = os.path.join(carpeta_copia)
        log_file_path = os.path.join(backup_folder, 'Network_host_log.txt')
        backup_imp = os.path.join(backup_folder, "Printers_log.txt")

        # Llamar a save_printer_log aquí
        self.save_printer_log(backup_imp)
        self.create_log(log_file_path)

        
    #ejecutamos el software SAP
    def restaurar_perfil(self):
        #preguntamos si el usuario quiere revertir una copia de seguridad con SAP
        reply = QtWidgets.QMessageBox.question(self, 'Ejecutar SAP', "¿Desea restaurar Conexiones SAP?", QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No, QtWidgets.QMessageBox.No)
        if reply == QtWidgets.QMessageBox.Yes:
            #ejecutamos el software SAP
            os.system("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe")
            time.sleep(10)
            os.system("taskkill /im saplogon.exe /f")
            self.restaurar_perfil_ok()
        else:
            self.restaurar_perfil_ok()
        



    def restaurar_perfil_ok(self):
        perfil_seleccionado = None
        for cb in self.checkbox_perfiles_restaurar:
            if cb.isChecked():
                perfil_seleccionado = cb.text()
                break
        if perfil_seleccionado is None:
            QtWidgets.QMessageBox.warning(self, "Error", "Por favor, seleccione un perfil para restaurar.")
            return

        zip_seleccionado = None
        for cb in self.checkbox_zips_restaurar:
            if cb.isChecked():
                zip_seleccionado = cb.text()
                break

        if zip_seleccionado is None:
            QtWidgets.QMessageBox.warning(self, "Error", "Por favor, seleccione un archivo ZIP para restaurar.")
            return

        carpeta_restaurar = self.label_5.text()
        if not os.path.exists(carpeta_restaurar):
            QtWidgets.QMessageBox.warning(self, "Error", "Por favor, seleccione una carpeta de restauración válida.")
            return

        ruta_usuarios = os.path.expanduser("~").rsplit(os.sep, 1)[0]
        ruta_perfil = os.path.join(ruta_usuarios, perfil_seleccionado)
        zip_path = os.path.join(carpeta_restaurar, zip_seleccionado)

        with zipfile.ZipFile(zip_path, 'r') as zf:
            zf.extractall(ruta_perfil)

        QtWidgets.QMessageBox.information(self, "Restauración realizada", "La restauración se realizó con éxito.")


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    SoftwareCopy = SoftwareCopyApp()
    SoftwareCopy.show()
    sys.exit(app.exec_())

