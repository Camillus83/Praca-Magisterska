
from PyQt5 import QtCore, QtGui, QtWidgets
import re, logging, math, openpyxl
from PyQt5.QtWidgets import QMessageBox
from pathlib import Path

logging.basicConfig(filename='programlogs.txt',filemode='w', level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s ')
logging.basicConfig()

class Ui_ekranStartowy(object):
    def menuprogramu(self):
        ekranStartowy.close()
        self.window2 = QtWidgets.QMainWindow()
        self.ui = Ui_menuProgramu()
        self.ui.setupUi(self.window2)
        self.window2.show()

    def setupUi(self, ekranStartowy):
        ekranStartowy.setObjectName("ekranStartowy")
        ekranStartowy.resize(516, 249)
        icon = QtGui.QIcon() 
        icon.addPixmap(QtGui.QPixmap(Path.cwd().as_posix()+"/icon.png"), QtGui.QIcon.Normal, QtGui.QIcon.On)
        ekranStartowy.setWindowIcon(icon)
        ekranStartowy.setWindowOpacity(1.0)
        self.centralwidget = QtWidgets.QWidget(ekranStartowy)
        self.centralwidget.setObjectName("centralwidget")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(40, 10, 491, 111))
        font = QtGui.QFont()
        font.setPointSize(14)
        font.setBold(True)
        font.setWeight(75)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(40, 110, 111, 21))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setGeometry(QtCore.QRect(40, 130, 141, 16))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_3.setFont(font)
        self.label_3.setObjectName("label_3")
        self.label_4 = QtWidgets.QLabel(self.centralwidget)
        self.label_4.setGeometry(QtCore.QRect(40, 150, 121, 16))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_4.setFont(font)
        self.label_4.setObjectName("label_4")
        self.label_5 = QtWidgets.QLabel(self.centralwidget)
        self.label_5.setGeometry(QtCore.QRect(40, 170, 121, 16))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_5.setFont(font)
        self.label_5.setObjectName("label_5")
        self.pushButton = QtWidgets.QPushButton(self.centralwidget, clicked = lambda:self.menuprogramu())
        self.pushButton.setGeometry(QtCore.QRect(260, 120, 161, 61))
        self.pushButton.setObjectName("pushButton")
        ekranStartowy.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(ekranStartowy)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 516, 21))
        self.menubar.setObjectName("menubar")
        ekranStartowy.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(ekranStartowy)
        self.statusbar.setObjectName("statusbar")
        ekranStartowy.setStatusBar(self.statusbar)

        self.retranslateUi(ekranStartowy)
        QtCore.QMetaObject.connectSlotsByName(ekranStartowy)

    def retranslateUi(self, ekranStartowy):
        _translate = QtCore.QCoreApplication.translate
        ekranStartowy.setWindowTitle(_translate("ekranStartowy", "Kalkulator instalacji NN"))
        self.label.setText(_translate("ekranStartowy", 
                                        "Kalkulator wspomagający projektowanie \n"
                                        "oprzewododowania oraz zabezpieczeń \n"
                                        "w instalacjach elektrycznych niskiego napięcia"))
        self.label_2.setText(_translate("ekranStartowy", "Kamil Richter"))
        self.label_3.setText(_translate("ekranStartowy", "Politechnika Białostocka"))
        self.label_4.setText(_translate("ekranStartowy", "Wydział Elektryczny"))
        self.label_5.setText(_translate("ekranStartowy", "2022"))
        self.pushButton.setText(_translate("ekranStartowy", "START"))

class Ui_menuProgramu(object):
    def setupUi(self, menuProgramu):
        menuProgramu.setObjectName("menuProgramu")
        menuProgramu.resize(793, 600)
        menuProgramu.setMinimumSize(QtCore.QSize(793, 600))
        menuProgramu.setMaximumSize(QtCore.QSize(793, 600))
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(Path.cwd().as_posix()+"/icon.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        menuProgramu.setWindowIcon(icon)
        self.centralwidget = QtWidgets.QWidget(menuProgramu)
        self.centralwidget.setMaximumSize(QtCore.QSize(16777215, 16777215))
        self.centralwidget.setObjectName("centralwidget")
        self.cosPhi_output = QtWidgets.QLabel(self.centralwidget)
        self.cosPhi_output.setGeometry(QtCore.QRect(370, 340, 141, 16))
        self.cosPhi_output.setObjectName("cosPhi_output")
        self.buttonZapiszPlik = QtWidgets.QPushButton(self.centralwidget, clicked=lambda:self.drukowanie())
        self.buttonZapiszPlik.setGeometry(QtCore.QRect(520, 500, 141, 51))
        self.buttonZapiszPlik.setObjectName("buttonZapiszPlik")
        self.Ib_output = QtWidgets.QLabel(self.centralwidget)
        self.Ib_output.setGeometry(QtCore.QRect(370, 300, 141, 16))
        self.Ib_output.setMouseTracking(False)
        self.Ib_output.setObjectName("Ib_output")
        self.Label_sposob_ulozenia = QtWidgets.QLabel(self.centralwidget)
        self.Label_sposob_ulozenia.setGeometry(QtCore.QRect(370, 140, 91, 16))
        self.Label_sposob_ulozenia.setObjectName("Label_sposob_ulozenia")
        self.Pc_output = QtWidgets.QLabel(self.centralwidget)
        self.Pc_output.setGeometry(QtCore.QRect(370, 320, 141, 16))
        self.Pc_output.setObjectName("Pc_output")
        self.napiecie_zasilania_Input = QtWidgets.QComboBox(self.centralwidget)
        self.napiecie_zasilania_Input.setGeometry(QtCore.QRect(710, 110, 51, 21))
        self.napiecie_zasilania_Input.setObjectName("napiecie_zasilania_Input")
        self.napiecie_zasilania_Input.addItem("")
        self.napiecie_zasilania_Input.addItem("")
        self.rodzaj_zasilania_Input = QtWidgets.QComboBox(self.centralwidget)
        self.rodzaj_zasilania_Input.setGeometry(QtCore.QRect(460, 110, 131, 22))
        self.rodzaj_zasilania_Input.setCursor(QtGui.QCursor(QtCore.Qt.ArrowCursor))
        self.rodzaj_zasilania_Input.setObjectName("rodzaj_zasilania_Input")
        self.rodzaj_zasilania_Input.addItem("")
        self.rodzaj_zasilania_Input.addItem("")
        self.Label_napiecie_zasilajace = QtWidgets.QLabel(self.centralwidget)
        self.Label_napiecie_zasilajace.setGeometry(QtCore.QRect(600, 110, 111, 16))
        self.Label_napiecie_zasilajace.setObjectName("Label_napiecie_zasilajace")
        self.Idd_jednozylowy_output = QtWidgets.QLabel(self.centralwidget)
        self.Idd_jednozylowy_output.setGeometry(QtCore.QRect(370, 400, 291, 16))
        self.Idd_jednozylowy_output.setObjectName("Idd_jednozylowy_output")
        self.buttonOblicz = QtWidgets.QPushButton(self.centralwidget, clicked=lambda:self.obliczanie())
        self.buttonOblicz.setGeometry(QtCore.QRect(370, 500, 141, 51))
        self.buttonOblicz.setObjectName("buttonOblicz")
        self.Label_dane_obciazenia = QtWidgets.QLabel(self.centralwidget)
        self.Label_dane_obciazenia.setGeometry(QtCore.QRect(60, 70, 251, 31))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.Label_dane_obciazenia.setFont(font)
        self.Label_dane_obciazenia.setObjectName("Label_dane_obciazenia")
        self.sposob_ulozenia_Input = QtWidgets.QComboBox(self.centralwidget)
        self.sposob_ulozenia_Input.setGeometry(QtCore.QRect(460, 140, 301, 22))
        self.sposob_ulozenia_Input.setObjectName("sposob_ulozenia_Input")
        self.sposob_ulozenia_Input.addItem("")
        self.sposob_ulozenia_Input.addItem("")
        self.sposob_ulozenia_Input.addItem("")
        self.Idd_wielozylowego_output = QtWidgets.QLabel(self.centralwidget)
        self.Idd_wielozylowego_output.setGeometry(QtCore.QRect(370, 454, 281, 16))
        self.Idd_wielozylowego_output.setObjectName("Idd_wielozylowego_output")
        self.S_wielozylowy_output = QtWidgets.QLabel(self.centralwidget)
        self.S_wielozylowy_output.setGeometry(QtCore.QRect(370, 420, 241, 16))
        self.S_wielozylowy_output.setObjectName("S_wielozylowy_output")
        self.S_jednozylowy_output = QtWidgets.QLabel(self.centralwidget)
        self.S_jednozylowy_output.setGeometry(QtCore.QRect(370, 380, 241, 16))
        self.S_jednozylowy_output.setObjectName("S_jednozylowy_output")
        self.Label_rodzaj_zasilania = QtWidgets.QLabel(self.centralwidget)
        self.Label_rodzaj_zasilania.setGeometry(QtCore.QRect(370, 110, 81, 16))
        self.Label_rodzaj_zasilania.setObjectName("Label_rodzaj_zasilania")
        self.Label_odleglosc = QtWidgets.QLabel(self.centralwidget)
        self.Label_odleglosc.setGeometry(QtCore.QRect(370, 170, 91, 16))
        self.Label_odleglosc.setObjectName("Label_odleglosc")
        self.odleglosc_Input = QtWidgets.QDoubleSpinBox(self.centralwidget)
        self.odleglosc_Input.setGeometry(QtCore.QRect(460, 170, 62, 22))
        self.odleglosc_Input.setMinimum(1.0)
        self.odleglosc_Input.setMaximum(1000.0)
        self.odleglosc_Input.setProperty("value", 1.0)
        self.odleglosc_Input.setObjectName("odleglosc_Input")
        self.Tytul_okna = QtWidgets.QLabel(self.centralwidget)
        self.Tytul_okna.setGeometry(QtCore.QRect(160, -10, 511, 61))
        font = QtGui.QFont()
        font.setPointSize(14)
        font.setBold(True)
        font.setWeight(75)
        self.Tytul_okna.setFont(font)
        self.Tytul_okna.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.Tytul_okna.setObjectName("Tytul_okna")
        self.Label_daneObwodu = QtWidgets.QLabel(self.centralwidget)
        self.Label_daneObwodu.setGeometry(QtCore.QRect(490, 80, 121, 16))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.Label_daneObwodu.setFont(font)
        self.Label_daneObwodu.setObjectName("Label_daneObwodu")
        self.DeltaU_jednozylowego_output = QtWidgets.QLabel(self.centralwidget)
        self.DeltaU_jednozylowego_output.setGeometry(QtCore.QRect(370, 360, 291, 16))
        self.DeltaU_jednozylowego_output.setObjectName("DeltaU_jednozylowego_output")
        self.Izabezpieczenia_output = QtWidgets.QLabel(self.centralwidget)
        self.Izabezpieczenia_output.setGeometry(QtCore.QRect(370, 470, 310, 16))
        self.Izabezpieczenia_output.setObjectName("Izabezpieczenia_output")
        self.tabelaObciazen = QtWidgets.QTableWidget(self.centralwidget)
        self.tabelaObciazen.setGeometry(QtCore.QRect(30, 110, 300, 341))
        self.tabelaObciazen.setMinimumSize(QtCore.QSize(281, 0))
        self.tabelaObciazen.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.tabelaObciazen.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.tabelaObciazen.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        self.tabelaObciazen.setObjectName("tabelaObciazen")
        self.tabelaObciazen.setColumnCount(2)
        self.tabelaObciazen.setRowCount(1)
        item = QtWidgets.QTableWidgetItem()
        self.tabelaObciazen.setVerticalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        item.setFont(font)
        self.tabelaObciazen.setHorizontalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        item.setFont(font)
        self.tabelaObciazen.setHorizontalHeaderItem(1, item)
        self.tabelaObciazen.horizontalHeader().setVisible(True)
        self.tabelaObciazen.horizontalHeader().setCascadingSectionResizes(False)
        self.tabelaObciazen.horizontalHeader().setDefaultSectionSize(142)
        self.tabelaObciazen.horizontalHeader().setMinimumSectionSize(39)
        self.buttonDodajObciazenie = QtWidgets.QPushButton(self.centralwidget, clicked=lambda:self.dodawanieObciazenia())
        self.buttonDodajObciazenie.setGeometry(QtCore.QRect(30, 460, 141, 23))
        self.buttonDodajObciazenie.setObjectName("buttonDodajObciazenie")
        self.buttonUsunObciazenie = QtWidgets.QPushButton(self.centralwidget, clicked=lambda:self.usuwanieObciazenia())
        self.buttonUsunObciazenie.setGeometry(QtCore.QRect(190, 460, 141, 23))
        self.buttonUsunObciazenie.setObjectName("buttonUsunObciazenie")
        self.Label_wynikiObliczen = QtWidgets.QLabel(self.centralwidget)
        self.Label_wynikiObliczen.setGeometry(QtCore.QRect(490, 275, 121, 21))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.Label_wynikiObliczen.setFont(font)
        self.Label_wynikiObliczen.setObjectName("Label_wynikiObliczen")
        self.Label_wsp_jednoczesnosci = QtWidgets.QLabel(self.centralwidget)
        self.Label_wsp_jednoczesnosci.setGeometry(QtCore.QRect(370, 200, 91, 31))
        self.Label_wsp_jednoczesnosci.setObjectName("Label_wsp_jednoczesnosci")
        self.wsp_jednoczesnosci_Input = QtWidgets.QDoubleSpinBox(self.centralwidget)
        self.wsp_jednoczesnosci_Input.setGeometry(QtCore.QRect(460, 200, 62, 31))
        self.wsp_jednoczesnosci_Input.setMaximum(1.0)
        self.wsp_jednoczesnosci_Input.setSingleStep(0.05)
        self.wsp_jednoczesnosci_Input.setProperty("value", 1.0)
        self.wsp_jednoczesnosci_Input.setObjectName("wsp_jednoczesnosci_Input")
        self.DeltaU_wielozylowego_output = QtWidgets.QLabel(self.centralwidget)
        self.DeltaU_wielozylowego_output.setGeometry(QtCore.QRect(370, 435, 291, 16))
        self.DeltaU_wielozylowego_output.setObjectName("DeltaU_wielozylowego_output")
        self.Label_Praca_Magisterska = QtWidgets.QLabel(self.centralwidget)
        self.Label_Praca_Magisterska.setGeometry(QtCore.QRect(190, 40, 411, 31))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.Label_Praca_Magisterska.setFont(font)
        self.Label_Praca_Magisterska.setObjectName("Label_Praca_Magisterska")
        self.Label_t_otoczenia = QtWidgets.QLabel(self.centralwidget)
        self.Label_t_otoczenia.setGeometry(QtCore.QRect(540, 200, 131, 31))
        self.Label_t_otoczenia.setObjectName("Label_t_otoczenia")
        self.Label_rodzaj_izolacji = QtWidgets.QLabel(self.centralwidget)
        self.Label_rodzaj_izolacji.setGeometry(QtCore.QRect(540, 170, 81, 16))
        self.Label_rodzaj_izolacji.setObjectName("Label_rodzaj_izolacji")
        self.rodzaj_izolacji_Input = QtWidgets.QComboBox(self.centralwidget)
        self.rodzaj_izolacji_Input.setGeometry(QtCore.QRect(620, 170, 131, 22))
        self.rodzaj_izolacji_Input.setObjectName("rodzaj_izolacji_Input")
        self.rodzaj_izolacji_Input.addItem("")
        self.rodzaj_izolacji_Input.addItem("")
        self.Label_rodzaj_zabezpieczenia = QtWidgets.QLabel(self.centralwidget)
        self.Label_rodzaj_zabezpieczenia.setGeometry(QtCore.QRect(370, 240, 111, 21))
        self.Label_rodzaj_zabezpieczenia.setObjectName("Label_rodzaj_zabezpieczenia")
        self.rodzaj_zabezpieczenia_Input = QtWidgets.QComboBox(self.centralwidget)
        self.rodzaj_zabezpieczenia_Input.setGeometry(QtCore.QRect(490, 240, 161, 21))
        self.rodzaj_zabezpieczenia_Input.setObjectName("rodzaj_zabezpieczenia_Input")
        self.rodzaj_zabezpieczenia_Input.addItem("")
        self.rodzaj_zabezpieczenia_Input.addItem("")
        self.t_otoczeniai_Input = QtWidgets.QDoubleSpinBox(self.centralwidget)
        self.t_otoczeniai_Input.setGeometry(QtCore.QRect(620, 200, 61, 31))
        self.t_otoczeniai_Input.setMinimum(-40.0)
        self.t_otoczeniai_Input.setMaximum(60.0)
        self.t_otoczeniai_Input.setSingleStep(5.0)
        self.t_otoczeniai_Input.setProperty("value", 30.0)
        self.t_otoczeniai_Input.setObjectName("t_otoczeniai_Input")

        self.rodzaj_izolacji_Input.activated.connect(self.zakresyTemperatur)
        
        menuProgramu.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(menuProgramu)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 770, 21))
        self.menubar.setObjectName("menubar")
        menuProgramu.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(menuProgramu)
        self.statusbar.setObjectName("statusbar")
        menuProgramu.setStatusBar(self.statusbar)

        self.tabelaObciazen.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.retranslateUi(menuProgramu)
        self.rodzaj_zasilania_Input.setCurrentIndex(0)
        self.sposob_ulozenia_Input.setCurrentIndex(0)
        QtCore.QMetaObject.connectSlotsByName(menuProgramu)
   
    # Zmiana ustawień zakresu temperatury otoczenia w zależnosci od rodzaju izolacji
    def zakresyTemperatur(self, index):
        self.t_otoczeniai_Input.clear()
        if index == 0: #POLWINIT PVC
            self.t_otoczeniai_Input.setMinimum(-40.0)
            self.t_otoczeniai_Input.setMaximum(60.0)
            self.t_otoczeniai_Input.setSingleStep(5.0)
            self.t_otoczeniai_Input.setProperty("value", 30.0)
        elif index == 1: #POLIETYLEN USIECIOWANY HALOGEN POWŁOKA
            self.t_otoczeniai_Input.setMinimum(-40.0)
            self.t_otoczeniai_Input.setMaximum(80.0)
            self.t_otoczeniai_Input.setSingleStep(5.0)
            self.t_otoczeniai_Input.setProperty("value", 30.0)
        elif index == 2:  #POLIETYLEN USIECIOWANY BEZHALOGENOWA POWŁOKA /EPR
            self.t_otoczeniai_Input.setMinimum(-40.0)
            self.t_otoczeniai_Input.setMaximum(60.0)
            self.t_otoczeniai_Input.setSingleStep(5.0)
            self.t_otoczeniai_Input.setProperty("value", 30.0)
        pass
    #Wyskakujące okienko w przypadku gdy wprowadzone dane są nieprawidłowe
    def BladDanychMessage(self):
        msg = QMessageBox()
        msg.setWindowTitle('Błąd danych wejściowych')
        msg.setText('Wystąpił błąd danych.\nCo najmniej jedna z danych wejściowych jest niepoprawna.')
        msg.setIcon(QMessageBox.Critical)
        msg.setDetailedText('Wartość mocy czynnej powinna być liczbą rzeczywistą wyrażoną w W. Np. 2000 \nWartość współczynnika mocy powinna być liczbą rzeczywistą mieszczącą się w zakresie między 0 a 1 włącznie. Np. 0.89')
        msg.exec()
    #Wyskakujące okienko w przypadku gdy niemożliwe jest spełnienie warunku obciążalności długotrwałej lub spadku napięcia
    def NiemozliwyWarunekMessage(self):
        msg = QMessageBox()
        msg.setWindowTitle('Błąd obliczeń')
        msg.setText('Nie można spełnić warunków dla wprowadzonych danych.')
        msg.setIcon(QMessageBox.Warning)
        msg.exec()
    #Dodawanie kolejnego obciążenia do tabeli odbiorników
    def dodawanieObciazenia(self):
        QtWidgets.QTableWidget.insertRow(self.tabelaObciazen, self.tabelaObciazen.rowCount())
    #Usuwanie zaznaczonego obciążenia z tabeli odbiorników
    def usuwanieObciazenia(self):
        if self.tabelaObciazen.currentRow() != 0:
            QtWidgets.QTableWidget.removeRow(self.tabelaObciazen, self.tabelaObciazen.currentRow())
        else:
            pass
    #Sprawdzanie czy wprowadzone dane są prawidłowe
    def sprawdzanieDanych(self):
        liczbyRegex = re.compile(r'^(\d*\.)?\d+$')
        flag = False
        logging.debug(f'Ilosc wierszy: {self.tabelaObciazen.rowCount()}')
        while True:
            for i in range(QtWidgets.QTableWidget.rowCount(self.tabelaObciazen)):
                for j in range(QtWidgets.QTableWidget.columnCount(self.tabelaObciazen)):
                    if self.tabelaObciazen.item(i,j) == None:
                        logging.debug('Wystapil blad pustej komorki!')
                        flag = True
                        break
                    logging.debug(f'Czy wystapil blad?: {flag}')
                    logging.debug(f'Zawartosc komorki: {self.tabelaObciazen.item(i,j).text()}')
                    testZawartosciKomorki = liczbyRegex.search(self.tabelaObciazen.item(i,j).text())
                    logging.debug(f'Wynik testu wyrazen regularnych komorki: {testZawartosciKomorki}')
                    if testZawartosciKomorki == None:
                        flag = True
                        logging.debug(f'Wystapil blad danych!')
                        break
                    if j == 1:
                        if float(self.tabelaObciazen.item(i,j).text()) > 1:
                            flag = True
                            logging.debug(f'Wystapil blad danych, wartosc wspolczynnika nie moze byc wieksza od 1')
                            break
                        break
            break
        if flag == True:
            logging.debug('DANE NIEPOPRAWNE')
        else:
            logging.debug('DANE POPRAWNE')
        return flag
    #Sprawdzanie czy prąd obliczeniowy jest mniejszy niż obciążalność długotrwała przewodu o największym przekroju
    def mozliwoscSpelnienia(self, prad, moc_czynna, odleglosc, napiecie, rodzaj_zasilania):
        flag = False
        if prad > float(tabelaObciazalnosciJednozylowych[list(tabelaObciazalnosciJednozylowych)[-1]]):
            flag = True
        if prad > float(tabelaObciazalnosciWielozylowych[list(tabelaObciazalnosciWielozylowych)[-1]]):
            flag = True
        if rodzaj_zasilania == 'Obwód jednofazowy':
            if (200*moc_czynna*odleglosc)/(56*float(list(tabelaObciazalnosciJednozylowych)[-1])*math.pow(napiecie,2)) > 3:
                flag = True
        if rodzaj_zasilania == "Obwód trójfazowy":
            if (100*moc_czynna*odleglosc)/(56*float(list(tabelaObciazalnosciJednozylowych)[-1])*math.pow(napiecie,2)) > 3:
                flag = True

        return flag
    #Odczytywanie danych dt. obciążalności długotrwałej przewodów z excela, stworzenie tabeli obciążalności i zabezpieczeń
    def odczytDanych(self, sposob_ulozenia, rodzaj_zasilania):
        wb = openpyxl.load_workbook('tabela_obciazalnosci2.xlsx')
        if sposob_ulozenia == "W rurkach i kanałach instalacyjnych pod tynkiem":
            sheet = wb['RK_pod_tynkiem']
        elif sposob_ulozenia == 'W rurkach i kanałach instalacyjnych na ścianie':
            sheet = wb['RK_na_scianie']
        elif sposob_ulozenia == 'Na ścianie':
            sheet = wb['Na_scianie']
        global tabelaObciazalnosciJednozylowych, tabelaObciazalnosciWielozylowych, typoszeregZabezpieczen
        #global t_gr, wsp_zabezpieczenia
        tabelaObciazalnosciJednozylowych = {}
        tabelaObciazalnosciWielozylowych = {}
        #typoszeregZabezpieczen = (16,20,25,35,40,63,80,100,125,160,200,250,315)
        rodzaj_izolacji = self.rodzaj_izolacji_Input.currentIndex()
        for i in range(4,sheet.max_row+1,1):
            if rodzaj_zasilania == "Obwód jednofazowy":
                if rodzaj_izolacji == 0:
                    tabelaObciazalnosciJednozylowych[sheet.cell(row=i, column=1).value] = sheet.cell(row=i, column=2).value
                    tabelaObciazalnosciWielozylowych[sheet.cell(row=i, column=1).value] = sheet.cell(row=i, column=6).value
                if rodzaj_izolacji == 1 or rodzaj_izolacji == 2:
                    tabelaObciazalnosciJednozylowych[sheet.cell(row=i, column=1).value] = sheet.cell(row=i, column=3).value
                    tabelaObciazalnosciWielozylowych[sheet.cell(row=i, column=1).value] = sheet.cell(row=i, column=7).value
            elif rodzaj_zasilania == 'Obwód trójfazowy':
                if rodzaj_izolacji == 0:
                    tabelaObciazalnosciJednozylowych[sheet.cell(row=i, column=1).value] = sheet.cell(row=i, column=4).value
                    tabelaObciazalnosciWielozylowych[sheet.cell(row=i, column=1).value] = sheet.cell(row=i, column=8).value
                elif rodzaj_izolacji == 1 or rodzaj_izolacji == 2:
                    tabelaObciazalnosciJednozylowych[sheet.cell(row=i, column=1).value] = sheet.cell(row=i, column=5).value
                    tabelaObciazalnosciWielozylowych[sheet.cell(row=i, column=1).value] = sheet.cell(row=i, column=9).value
        logging.debug('Tabela obciazalnosci przewodow jednozylowych: ' + str(tabelaObciazalnosciJednozylowych))
        logging.debug('Tabela obciazalnosci przewodow wielozylowych: ' + str(tabelaObciazalnosciWielozylowych))
        
    #Funkcja uwzględniająca temperaturę otoczenia, zmienia wartość obciążalności długotrwałej przewodu.
    def UwzglednienieTemperaturyOtoczenia(self, t_graniczna, temperatura_otoczenia, tabelaObciazalnosciJednozylowych, tabelaObciazalnosciWielozylowych):
        temperatura_otoczenia_nominalna = 30
        #Temperatura graniczna dopuszczalna
        t_gr_dop = t_graniczna
        for key in tabelaObciazalnosciWielozylowych:
            tabelaObciazalnosciJednozylowych[key] = round(tabelaObciazalnosciJednozylowych[key] * math.sqrt((t_gr_dop - temperatura_otoczenia)/(t_gr_dop - temperatura_otoczenia_nominalna)),2) 
            tabelaObciazalnosciWielozylowych[key] = round(tabelaObciazalnosciWielozylowych[key] * math.sqrt((t_gr_dop - temperatura_otoczenia)/(t_gr_dop - temperatura_otoczenia_nominalna)),2)
            
        logging.debug('Tabela obciazalnosci przewodow jednozylowych po uwzględnieniu temperatury otoczenia: ' + str(tabelaObciazalnosciJednozylowych))
        logging.debug('Tabela obciazalnosci przewodow wielozylowych po uwzględnieniu temperatury otoczenia: ' + str(tabelaObciazalnosciWielozylowych))
   
    #Funkcja która bierze kolejny przekrój ze słownika
    def kolejnyPrzekroj(self, test_dict, test_key):
        res = None
        temp = iter(test_dict)
        for key in temp:
            if key == test_key:
                res = next(temp, key)
        return res
    #Funkcja sprawdzajaca warunek obciążalności długotrwałej przewodu
    def sprawdzWarunek1(self, aktualnyPrzekroj, pradObliczeniowyObwodu):#,tabelaObciazalnosciJednozylowych, tabelaObciazalnosciWielozylowych):
        while True:
            logging.debug(f'Prad obliczeniowy obwodu wynosi: {round(pradObliczeniowyObwodu,2)}')
            logging.debug(f'Sprawdzanie warunku obciazalnosci dlugotrwalej dla przewodu jednozylowego o przekroju {aktualnyPrzekroj[0]} mm2 oraz przewodu wielozylowego o przekroju {aktualnyPrzekroj[1]} mm2')
            if pradObliczeniowyObwodu < tabelaObciazalnosciJednozylowych[aktualnyPrzekroj[0]] and pradObliczeniowyObwodu < tabelaObciazalnosciWielozylowych[aktualnyPrzekroj[1]]:
                logging.debug(f'Warunek obciazalnosci dlugotrwalej jest spelniony dla przewodu jednozylowego o przekroju {aktualnyPrzekroj[0]} mm2 o obciazalnosci dlugotrwalej rownej {tabelaObciazalnosciJednozylowych[aktualnyPrzekroj[0]]} A')
                logging.debug(f'Warunek obciazalnosci dlugotrwalej jest spelniony dla przewodu wielozylowego o przekroju {aktualnyPrzekroj[1]} mm2 o obciazalnosci dlugotrwalej rownej {tabelaObciazalnosciWielozylowych[aktualnyPrzekroj[1]]} A')
                break
            if pradObliczeniowyObwodu > tabelaObciazalnosciJednozylowych[aktualnyPrzekroj[0]]:
                logging.debug(f'Obciazalnosc pradowa przewodu jednozylowego {tabelaObciazalnosciJednozylowych[aktualnyPrzekroj[0]]} A o przekroju {aktualnyPrzekroj} mm2 jest za mala')
                aktualnyPrzekroj[0] = self.kolejnyPrzekroj(tabelaObciazalnosciJednozylowych, aktualnyPrzekroj[0])
            if pradObliczeniowyObwodu > tabelaObciazalnosciWielozylowych[aktualnyPrzekroj[1]]:
                logging.debug(f'Obciazalnosc pradowa przewodu wielozylowego {tabelaObciazalnosciWielozylowych[aktualnyPrzekroj[1]]} A dla przekroju {aktualnyPrzekroj[1]} mm2 jest za mala')
                aktualnyPrzekroj[1] = self.kolejnyPrzekroj(tabelaObciazalnosciJednozylowych, aktualnyPrzekroj[1])
        return aktualnyPrzekroj
    #Funkcja sprawdzająca warunek dopuszczalnego spadku napięcia
    def sprawdzWarunek2(self, aktualnyPrzekroj, mocCalkowitaObwodu, odleglosc, rodzaj_zasilania, napiecie_zasilajace):
        global spadek_napieciaJednozylowy
        global spadek_napieciaWielozylowy
        while True:
            logging.debug(f'{aktualnyPrzekroj}')
            logging.debug(f'{tabelaObciazalnosciJednozylowych}')
            logging.debug(f'Rodzaj zasilania: {rodzaj_zasilania} A, Aktualny przekroj przewodu jednozylowego: {aktualnyPrzekroj[0]} mm2, Aktualny przekroj przewodu wielozylowego: {aktualnyPrzekroj[1]} mm2, Napiecie zasilajace: {napiecie_zasilajace} V')
            if rodzaj_zasilania == 'Obwód jednofazowy':
                logging.debug(f'Moc calkowita obwodu: {mocCalkowitaObwodu}, odleglosc: {odleglosc}, aktualny przekroj: {aktualnyPrzekroj[0]} Napiecie zasilajace: {napiecie_zasilajace} ')
                spadek_napieciaJednozylowy = ((200*mocCalkowitaObwodu * odleglosc)/(56*aktualnyPrzekroj[0]*pow(napiecie_zasilajace,2)))
                spadek_napieciaWielozylowy = ((200*mocCalkowitaObwodu * odleglosc)/(56*aktualnyPrzekroj[1]*pow(napiecie_zasilajace,2)))
            if rodzaj_zasilania == 'Obwód trójfazowy':
                spadek_napieciaJednozylowy = ((100*mocCalkowitaObwodu * odleglosc)/(56*aktualnyPrzekroj[0]*pow(napiecie_zasilajace,2)))
                spadek_napieciaWielozylowy = ((100*mocCalkowitaObwodu * odleglosc)/(56*aktualnyPrzekroj[1]*pow(napiecie_zasilajace,2)))
            if spadek_napieciaJednozylowy > 3:
                logging.debug(f'Spadek napiecia dla przewodu jednozylowego o przekroju {aktualnyPrzekroj[0]} mm2 jest wiekszy od 3% i wynosi {round(spadek_napieciaJednozylowy,2)} %')
                aktualnyPrzekroj[0] = self.kolejnyPrzekroj(tabelaObciazalnosciJednozylowych, aktualnyPrzekroj[0])
            if spadek_napieciaWielozylowy > 3:
                logging.debug(f'Spadek napiecia dla przewodu wielozylowego o przekroju {aktualnyPrzekroj[1]} mm2 jest wiekszy od 3% i wynosi {round(spadek_napieciaWielozylowy,2)} %')
                aktualnyPrzekroj[1] = self.kolejnyPrzekroj(tabelaObciazalnosciWielozylowych, aktualnyPrzekroj[1])

            elif spadek_napieciaJednozylowy < 3 and spadek_napieciaWielozylowy < 3 :
                logging.debug(f'Warunek dotyczacy dopuszczalnego spadku napiecia jest spelniony dla przewodu jednozylowego o przekroju {aktualnyPrzekroj[0]} mm2 a spadek napiecia wynosi {round(spadek_napieciaJednozylowy,2)} %')
                break
        return aktualnyPrzekroj

 
    #Wyskakujące okienko gdy zabezpieczenie jest złe       
    def Bladzabezpieczenia(self, rodzaj):
        msg = QMessageBox()
        msg.setWindowTitle('CHUJOWE ZABEZPIECZENIE')
        msg.setText(f'Warunek niemożliwy do spełnienia w przypadku\n przewodu {rodzaj}')
        msg.setIcon(QMessageBox.Warning)
        msg.exec()
    #Główna funkcja licząca
    def sprawdzZabezpieczenie(self, wsp_zabezpieczenia, aktualnyPrzekroj, pradObliczeniowyObwodu, aktualneZabezpieczenie):
        wspolczynnik_krotnosci_pradu = wsp_zabezpieczenia
        i_jednozylowe = typoszeregZabezpieczen.index(aktualneZabezpieczenie[0])
        i_wielozylowe = typoszeregZabezpieczen.index(aktualneZabezpieczenie[1])
       #sprawdzanie warunku pierwszego
        logging.debug(f'START FUNKCJI DOBIERANIA WARTOŚCI ZABEZPIECZENIA...')
        logging.debug(f'Wartość prądu obliczeniowego: {round(pradObliczeniowyObwodu,2)} A')
        logging.debug(f'Wartość przekroju przewodu: {aktualnyPrzekroj[0]} mm2')
        logging.debug(f'Wartość obciazalności długotrwałej przewodu: {tabelaObciazalnosciJednozylowych[aktualnyPrzekroj[0]]} A')
        

        logging.debug(f'*************************************************************')
        logging.debug(f'SPRAWDZANIE WARUNKU 1')
        while True:
            # Sprawdzanie warunku 1 dla przewodów jednożyłowych
            if pradObliczeniowyObwodu > aktualneZabezpieczenie[0] or tabelaObciazalnosciJednozylowych[aktualnyPrzekroj[0]] < aktualneZabezpieczenie[0]:
                i_jednozylowe+=1
                logging.debug('Warunek 1 doboru zabezpieczeń niespełniony dla przewodu jednożyłowego')
                logging.debug(f'Aktualna wartość zabezpieczenia dla przewodów jednożyłowych: {aktualneZabezpieczenie[0]} A')
                if i_jednozylowe == len(typoszeregZabezpieczen):
                    aktualnyPrzekroj[0] = self.kolejnyPrzekroj(tabelaObciazalnosciJednozylowych, aktualnyPrzekroj[0])
                    i_jednozylowe=0
                    logging.debug('Ze względu na dobór zabezpieczenia przekrój przewodu został zwiększony')
                    logging.debug(f'Aktualny przekrój przewodu jednożyłowego wynosi: {aktualnyPrzekroj[0]} mm2')                       
                logging.debug(f'*************************************************************')
                aktualneZabezpieczenie[0] = typoszeregZabezpieczen[i_jednozylowe]

            # Sprawdzanie warunku 1 dla przewodów wielożyłowych
            if pradObliczeniowyObwodu > aktualneZabezpieczenie[1] or tabelaObciazalnosciWielozylowych[aktualnyPrzekroj[1]] < aktualneZabezpieczenie[1]:
                i_wielozylowe+=1
                logging.debug('Warunek 1 doboru zabezpieczeń niespełniony dla przewodu wielożyłowego')
                logging.debug(f'Aktualna wartość zabezpieczenia dla przewodów wielożyłowych: {aktualneZabezpieczenie[1]} A')
                if i_wielozylowe == len(typoszeregZabezpieczen):
                    aktualnyPrzekroj[1] = self.kolejnyPrzekroj(tabelaObciazalnosciWielozylowych, aktualnyPrzekroj[1])
                    i_wielozylowe=0
                    logging.debug('Ze względu na dobór zabezpieczenia przekrój przewodu wielożylowego został zwiększony')
                    logging.debug(f'Aktualny przekrój przewodu wielożyłowego wynosi: {aktualnyPrzekroj[1]} mm2') 
                logging.debug(f'*************************************************************')
                aktualneZabezpieczenie[1] = typoszeregZabezpieczen[i_wielozylowe]
            
            # Warunek 1 spełniony dla przewodów jednożyłowych i wielożyłowych
            elif pradObliczeniowyObwodu <= aktualneZabezpieczenie[0] and tabelaObciazalnosciJednozylowych[aktualnyPrzekroj[0]] >= aktualneZabezpieczenie[0] and pradObliczeniowyObwodu <= aktualneZabezpieczenie[1] and tabelaObciazalnosciWielozylowych[aktualnyPrzekroj[1]] >= aktualneZabezpieczenie[1]:
                logging.debug(f'Aktualna wartość zabezpieczenia dla przewodu jednożyłowego: {aktualneZabezpieczenie[0]} A')
                logging.debug(f'Aktualna wartość zabezpieczenia dla przewodu wielożyłowego: {aktualneZabezpieczenie[1]} A')
                logging.debug('Warunek 1 doboru zabezpieczeń spełniony')
                logging.debug(f'*************************************************************')
                break

        logging.debug(f'SPRAWDZANIE WARUNKU 2')
        logging.debug(f'Wartość prądu obliczeniowego: {round(pradObliczeniowyObwodu,2)} A')
        logging.debug(f'Wartość przekroju przewodu jednożyłowego: {aktualnyPrzekroj[0]} mm2')
        logging.debug(f'Wartość obciazalności długotrwałej przewodu jednożyłowego: {tabelaObciazalnosciJednozylowych[aktualnyPrzekroj[0]]} A')
        logging.debug(f'Aktualna wartość zabezpieczenia dla przewodu jednożyłowego: {aktualneZabezpieczenie[0]} A')
        while True:
            #Sprawdzanie warunku 2 dla przewodów jednożyłowych
            if aktualneZabezpieczenie[0] * wspolczynnik_krotnosci_pradu > 1.45 * tabelaObciazalnosciJednozylowych[aktualnyPrzekroj[0]]:
                logging.debug('Warunek 2 doboru zabezpieczeń niespełniony dla przewodu jednożyłowego')
                logging.debug(f'Aktualna wartość zabezpieczenia dla przewodów jednożyłowych: {aktualneZabezpieczenie[0]} A')
                logging.debug(f'I2 = {round(wspolczynnik_krotnosci_pradu*aktualneZabezpieczenie[0],2)} > 1,45*Idd = {round(1.45*tabelaObciazalnosciJednozylowych[aktualnyPrzekroj[0]],2)}')
                aktualnyPrzekroj[0] = self.kolejnyPrzekroj(tabelaObciazalnosciJednozylowych, aktualnyPrzekroj[0])
            #Sprawdzanie warunku 2 dla przewodów wielożyłowych
            if aktualneZabezpieczenie[1] * wspolczynnik_krotnosci_pradu > 1.45 * tabelaObciazalnosciWielozylowych[aktualnyPrzekroj[1]]:
                logging.debug('Warunek 2 doboru zabezpieczeń niespełniony dla przewodu wielożyłowego')
                logging.debug(f'Aktualna wartość zabezpieczenia dla przewodów wielożyłowych: {aktualneZabezpieczenie[1]} A')
                logging.debug(f'I2 = {round(wspolczynnik_krotnosci_pradu*aktualneZabezpieczenie[1],2)} > 1,45*Idd = {round(1.45*tabelaObciazalnosciWielozylowych[aktualnyPrzekroj[1]],2)}')
                aktualnyPrzekroj[1] = self.kolejnyPrzekroj(tabelaObciazalnosciWielozylowych, aktualnyPrzekroj[1])
            # Warunek 2 spełniony dla przewodów jednożyłowych i wielożyłowych
            if aktualneZabezpieczenie[0] * wspolczynnik_krotnosci_pradu <= 1.45 * tabelaObciazalnosciJednozylowych[aktualnyPrzekroj[0]] and aktualneZabezpieczenie[1] * wspolczynnik_krotnosci_pradu <= 1.45 * tabelaObciazalnosciWielozylowych[aktualnyPrzekroj[1]]:
                logging.debug('Warunek 2 doboru zabezpieczeń spełniony')
                logging.debug(f'Aktualna wartość zabezpieczenia dla przewodów jednożyłowych: {aktualneZabezpieczenie[0]}')
                logging.debug(f'I2 = {round(wspolczynnik_krotnosci_pradu*aktualneZabezpieczenie[0],2)} <= 1,45*Idd = {round(1.45*tabelaObciazalnosciJednozylowych[aktualnyPrzekroj[0]],2)}')
                logging.debug(f'Aktualna wartość zabezpieczenia dla przewodów wielożyłowych: {aktualneZabezpieczenie[1]} A')
                logging.debug(f'I2 = {round(wspolczynnik_krotnosci_pradu*aktualneZabezpieczenie[1],2)} <= 1,45*Idd = {round(1.45*tabelaObciazalnosciWielozylowych[aktualnyPrzekroj[1]],2)}')
                logging.debug(f'*************************************************************')
                break
        
        return aktualneZabezpieczenie  
    #Główna funkcja licząca
    def obliczanie(self):
        rodzaj_zasilania = self.rodzaj_zasilania_Input.currentText()
        napiecie_zasilajace = float(self.napiecie_zasilania_Input.currentText())
        sposob_ulozenia = self.sposob_ulozenia_Input.currentText()
        odleglosc = self.odleglosc_Input.value()
        rodzaj_izolacji = self.rodzaj_izolacji_Input.currentIndex()
        rodzaj_zabezpieczenia = self.rodzaj_zabezpieczenia_Input.currentIndex()
        t_otoczenia = self.t_otoczeniai_Input.value()
        global typoszeregZabezpieczen
        typoszeregZabezpieczen = (16,20,25,35,40,63,80,100,125,160,200,250,315)
        
        if rodzaj_zabezpieczenia == 0:
            wsp_zabezpieczenia = 1.45
        elif rodzaj_zabezpieczenia == 1:
            wsp_zabezpieczenia = 1.60
        if rodzaj_izolacji == 0:
            t_gr = 70
        elif rodzaj_izolacji == 1 or rodzaj_izolacji == 2:
            t_gr = 105

        logging.debug('Wybrany rodzaj zabezpieczenia: ' + str(self.rodzaj_zabezpieczenia_Input.currentText()))
        logging.debug('Współczynnik zabezpieczenia: ' + str(wsp_zabezpieczenia))
       
        Pc = 0
        Sc = 0
        Ib = 0
        CosObwodu = 0
        global aktualneZabezpieczenie
        while True:
            if self.sprawdzanieDanych() == True:
                self.BladDanychMessage()
                break
            for i in range(QtWidgets.QTableWidget.rowCount(self.tabelaObciazen)):
                Pc += float(self.tabelaObciazen.item(i,0).text()) * self.wsp_jednoczesnosci_Input.value()
                Sc += float(self.tabelaObciazen.item(i,0).text()) / float(self.tabelaObciazen.item(i,1).text()) * self.wsp_jednoczesnosci_Input.value()
            CosObwodu = Pc / Sc
            if rodzaj_zasilania == "Obwód jednofazowy":
                Ib = Pc / (napiecie_zasilajace*CosObwodu) 
            elif rodzaj_zasilania == "Obwód trójfazowy":
                Ib = Pc / (napiecie_zasilajace*CosObwodu*math.sqrt(3))

            self.odczytDanych(sposob_ulozenia, rodzaj_zasilania)
            self.UwzglednienieTemperaturyOtoczenia(t_gr, t_otoczenia, tabelaObciazalnosciJednozylowych,tabelaObciazalnosciWielozylowych)
            if self.mozliwoscSpelnienia(Ib, Pc, odleglosc, napiecie_zasilajace, rodzaj_zasilania) == True:
                self.NiemozliwyWarunekMessage()
                break
            aktualnyPrzekroj = [1.5, 1.5]
            aktualneZabezpieczenie = [16.0, 16.0]
            aktualnyPrzekroj = self.sprawdzWarunek1(aktualnyPrzekroj, Ib)
            aktualnyPrzekroj = self.sprawdzWarunek2(aktualnyPrzekroj, Pc, odleglosc, rodzaj_zasilania, napiecie_zasilajace)
            aktualneZabezpieczenie = self.sprawdzZabezpieczenie(wsp_zabezpieczenia, aktualnyPrzekroj, Ib, aktualneZabezpieczenie)
            self.WyswietlWyniki(Ib, Pc, CosObwodu, sposob_ulozenia, aktualnyPrzekroj)
            break
    
    def WyswietlWyniki(self, Ib, Pc, CosObwodu, sposob_ulozenia, aktualnyPrzekroj):
            self.Ib_output.setText(f'Prąd obliczeniowy obwodu: {round(Ib,2)} A')
            self.Ib_output.adjustSize()
            self.Pc_output.setText(f'Moc czynna obwodu: {round(Pc,2)} W')
            self.Pc_output.adjustSize()
            self.cosPhi_output.setText(f'Współczynnik mocy obwodu: {round(CosObwodu,2)}')
            self.cosPhi_output.adjustSize()
            if sposob_ulozenia == 'Na ścianie':
                self.S_jednozylowy_output.setText(f'Obliczeniowy przekrój przewodu jednożyłowego: - ')
                self.Idd_jednozylowy_output.setText(f'Obciążalność długotrwała przewodu: - ')
                self.DeltaU_jednozylowego_output.setText('Procentowy spadek napięcia dla przewodu jednożyłowego: -')
            else:
                self.S_jednozylowy_output.setText(f'Obliczeniowy przekrój przewodu jednożyłowego: {aktualnyPrzekroj[0]} mm2')
                self.Idd_jednozylowy_output.setText(f'Obciążalność długotrwała przewodu: {tabelaObciazalnosciJednozylowych[aktualnyPrzekroj[0]]} A')
            self.S_jednozylowy_output.adjustSize()
            self.S_wielozylowy_output.setText(f'Obliczeniowy przekrój przewodu wielożyłowego: {aktualnyPrzekroj[1]} mm2')
            self.S_wielozylowy_output.adjustSize()
            self.Idd_jednozylowy_output.adjustSize()
            self.Idd_wielozylowego_output.setText(f'Obciążalność długotrwała przewodu: {tabelaObciazalnosciWielozylowych[aktualnyPrzekroj[1]]} A')
            self.Idd_wielozylowego_output.adjustSize()
            self.Izabezpieczenia_output.setText(f"Sugerowany prąd znamionowy zabezpieczenia przetężeniowego: {aktualneZabezpieczenie[0]} / {aktualneZabezpieczenie[1]} A")
            self.Izabezpieczenia_output.adjustSize()
            self.DeltaU_jednozylowego_output.setText(f'Procentowy spadek napięcia dla przewodu jednożyłowego: {round(spadek_napieciaJednozylowy,2)} %')
            self.DeltaU_jednozylowego_output.adjustSize()
            self.DeltaU_wielozylowego_output.setText(f'Procentowy spadek napięcia dla przewodu wielożyłowego: {round(spadek_napieciaWielozylowy,2)} %')
            self.DeltaU_wielozylowego_output.adjustSize()
    
    #Funkcja drukująca wyniki do pliku .txt
    def drukowanie(self):
        plikWynikowy = open('wyniki.txt','w', encoding="utf-8") #checkthis
        plikWynikowy.write('----------------------------------------------------------------------------------------------------\n')
        plikWynikowy.write('| 			Kalkulator przekroju przewodów i zabezpieczeń                              |\n')
        plikWynikowy.write('|--------------------------------------------------------------------------------------------------|\n')
        plikWynikowy.write('| 				Dane poszczególnych odbiorników:		  		   |\n')
        plikWynikowy.write('|--------------------------------------------------------------------------------------------------|\n')
        for i in range(QtWidgets.QTableWidget.rowCount(self.tabelaObciazen)):
            plikWynikowy.write('|')
            obciazenie = 'Obciazenie nr ' + str(i+1)
            x = obciazenie.center(27,' ')
            plikWynikowy.write(x)
            plikWynikowy.write('|')
            mocCzynna = 'Moc czynna: ' + str(self.tabelaObciazen.item(i,0).text()) + ' W'
            x = mocCzynna.center(34,' ')
            plikWynikowy.write(x)
            plikWynikowy.write('|')
            wspolczynnikMocy = 'Współczynnik mocy: ' + str(self.tabelaObciazen.item(i,1).text())
            x = wspolczynnikMocy.center(35,' ')
            plikWynikowy.write(x)
            plikWynikowy.write('|\n')
        plikWynikowy.write('|--------------------------------------------------------------------------------------------------|\n')
        plikWynikowy.write('|')
        x = 'Rodzaj zasilania: ' + str(self.rodzaj_zasilania_Input.currentText())
        y = x.center(40, ' ')
        plikWynikowy.write(y)
        plikWynikowy.write('|')
        x = 'Napięcie zasilania: ' + str(self.napiecie_zasilania_Input.currentText()) + ' V'
        y = x.center(57,' ')
        plikWynikowy.write(y)
        plikWynikowy.write('|\n')
        plikWynikowy.write('|--------------------------------------------------------------------------------------------------|\n')
        plikWynikowy.write('|')
        x = 'Sposób ułożenia: ' + str(self.sposob_ulozenia_Input.currentText())
        y = x.center(98,' ')
        plikWynikowy.write(y)
        plikWynikowy.write('|\n')
        plikWynikowy.write('|--------------------------------------------------------------------------------------------------|\n')
        plikWynikowy.write('|')
        x = 'Odległość: ' + str(self.odleglosc_Input.value()) + ' m'
        y = x.center(57,' ')
        plikWynikowy.write(y)
        plikWynikowy.write('|')
        x = 'Współczynnik jednoczesności: ' + str(round(self.wsp_jednoczesnosci_Input.value(),2))
        y = x.center(40,' ')
        plikWynikowy.write(y)
        plikWynikowy.write('|\n')
        plikWynikowy.write('|--------------------------------------------------------------------------------------------------|\n')
        plikWynikowy.write('|')
        x = 'Temperatura otoczenia: ' + str(self.t_otoczeniai_Input.value()) + ' C'
        y = x.center(40,' ')
        plikWynikowy.write(y)
        plikWynikowy.write('|')
        x = 'Rodzaj izolacji: ' + str(self.rodzaj_izolacji_Input.currentText())
        y = x.center(57,' ')
        plikWynikowy.write(y)
        plikWynikowy.write('|\n')
        plikWynikowy.write('|--------------------------------------------------------------------------------------------------|\n')


        plikWynikowy.write('| 					Wyniki obliczeń:      	    		                   |\n')
        plikWynikowy.write('|--------------------------------------------------------------------------------------------------|\n')
        plikWynikowy.write('|')
        x = str(self.Ib_output.text())
        y = x.center(33,' ')
        plikWynikowy.write(y)
        plikWynikowy.write('|')
        x = str(self.Pc_output.text())
        y = x.center(29,' ')
        plikWynikowy.write(y)
        plikWynikowy.write('|')
        x = str(self.cosPhi_output.text())
        y = x.center(34,' ')
        plikWynikowy.write(y)
        plikWynikowy.write('|\n')
        plikWynikowy.write('|--------------------------------------------------------------------------------------------------|\n')
        plikWynikowy.write('|')
        x = str(self.S_jednozylowy_output.text())
        y = x.center(98,' ')
        plikWynikowy.write(y)
        plikWynikowy.write('|\n')
        plikWynikowy.write('|--------------------------------------------------------------------------------------------------|\n')
        plikWynikowy.write('|')
        x = str(self.DeltaU_jednozylowego_output.text())
        y = x.center(98,' ')
        plikWynikowy.write(y)
        plikWynikowy.write('|\n')
        plikWynikowy.write('|--------------------------------------------------------------------------------------------------|\n')
        plikWynikowy.write('|')
        x = str(self.Idd_jednozylowy_output.text())
        y = x.center(58,' ')    
        plikWynikowy.write(y)
        plikWynikowy.write('|')
        x = 'Prąd znamionowy zabezpieczenia: ' + str(aktualneZabezpieczenie[0]) +' A'
        y = x.center(39,' ')
        plikWynikowy.write(y)
        plikWynikowy.write('|\n')
        plikWynikowy.write('|--------------------------------------------------------------------------------------------------|\n')
        plikWynikowy.write('|')
        x = str(self.S_wielozylowy_output.text())
        y = x.center(98,' ')
        plikWynikowy.write(y)
        plikWynikowy.write('|\n')
        plikWynikowy.write('|')
        x = str(self.DeltaU_wielozylowego_output.text())
        y = x.center(98,' ')
        plikWynikowy.write(y)
        plikWynikowy.write('|\n')
        plikWynikowy.write('|--------------------------------------------------------------------------------------------------|\n')
        plikWynikowy.write('|')
        x = str(self.Idd_wielozylowego_output.text())
        y = x.center(58,' ')    
        plikWynikowy.write(y)
        plikWynikowy.write('|')
        x = 'Prąd znamionowy zabezpieczenia: ' + str(aktualneZabezpieczenie[1]) + ' A'
        y = x.center(39,' ')
        plikWynikowy.write(y)
        plikWynikowy.write('|\n')
        plikWynikowy.write('|--------------------------------------------------------------------------------------------------|\n')
    
    
    def retranslateUi(self, menuProgramu):
        _translate = QtCore.QCoreApplication.translate
        menuProgramu.setWindowTitle(_translate("menuProgramu", "Kalkulator instalacji NN"))
        self.cosPhi_output.setText(_translate("menuProgramu", "Współczynnik mocy obwodu:"))
        self.buttonZapiszPlik.setText(_translate("menuProgramu", "Zapisz dane do pliku"))
        self.Ib_output.setText(_translate("menuProgramu", "Prąd obliczeniowy obwodu:"))
        self.Label_sposob_ulozenia.setText(_translate("menuProgramu", "Sposób ułożenia:"))
        self.Pc_output.setText(_translate("menuProgramu", "Moc czynna obwodu:"))
        self.napiecie_zasilania_Input.setItemText(0, _translate("menuProgramu", "230"))
        self.napiecie_zasilania_Input.setItemText(1, _translate("menuProgramu", "400"))
        self.rodzaj_zasilania_Input.setItemText(0, _translate("menuProgramu", "Obwód jednofazowy"))
        self.rodzaj_zasilania_Input.setItemText(1, _translate("menuProgramu", "Obwód trójfazowy"))
        self.Label_napiecie_zasilajace.setText(_translate("menuProgramu", "Napięcie fazowe [V]"))
        self.Idd_jednozylowy_output.setText(_translate("menuProgramu", "Obciążalność długotrwała przewodu jednożyłowego:"))
        self.buttonOblicz.setText(_translate("menuProgramu", "Oblicz"))
        self.Label_dane_obciazenia.setText(_translate("menuProgramu", "Dane poszczególnych odbiorników"))
        self.sposob_ulozenia_Input.setItemText(0, _translate("menuProgramu", "W rurkach i kanałach instalacyjnych pod tynkiem"))
        self.sposob_ulozenia_Input.setItemText(1, _translate("menuProgramu", "W rurkach i kanałach instalacyjnych na ścianie"))
        self.sposob_ulozenia_Input.setItemText(2, _translate("menuProgramu", "Na ścianie"))
        self.Idd_wielozylowego_output.setText(_translate("menuProgramu", "Obciążalność długotrwała przewodu wielożyłowego:"))
        self.S_wielozylowy_output.setText(_translate("menuProgramu", "Obliczeniowy przekrój przewodu wielożyłowego:"))
        self.S_jednozylowy_output.setText(_translate("menuProgramu", "Obliczeniowy przekrój przewodu jednożyłowego:"))
        self.Label_rodzaj_zasilania.setText(_translate("menuProgramu", "Rodzaj zasilania:"))
        self.Label_odleglosc.setText(_translate("menuProgramu", "Odległość: [m]"))
        self.Tytul_okna.setText(_translate("menuProgramu", "Kalkulator przekroju przewodów i zabezpieczeń"))
        self.Label_daneObwodu.setText(_translate("menuProgramu", "Dane obwodu"))
        self.DeltaU_jednozylowego_output.setText(_translate("menuProgramu", "Procentowy spadek napięcia dla przewodu jednożyłowego:"))
        self.Izabezpieczenia_output.setText(_translate("menuProgramu", "Sugerowany prąd znamionowy zabezpieczenia przetężeniowego:"))
        item = self.tabelaObciazen.verticalHeaderItem(0)
        item.setText(_translate("menuProgramu", "1"))
        item = self.tabelaObciazen.horizontalHeaderItem(0)
        item.setText(_translate("menuProgramu", "Moc czynna [W]"))
        item = self.tabelaObciazen.horizontalHeaderItem(1)
        item.setText(_translate("menuProgramu", "Współczynnik mocy "))
        self.buttonDodajObciazenie.setText(_translate("menuProgramu", "Dodaj odbiornik"))
        self.buttonUsunObciazenie.setText(_translate("menuProgramu", "Usuń odbiornik"))
        self.Label_wynikiObliczen.setText(_translate("menuProgramu", "Wyniki obliczeń"))
        self.Label_wsp_jednoczesnosci.setText(_translate("menuProgramu", "Współczynnik\n"
"jednoczesności"))
        self.DeltaU_wielozylowego_output.setText(_translate("menuProgramu", "Procentowy spadek napięcia dla przewodu wielożyłowego:"))
        self.Label_Praca_Magisterska.setText(_translate("menuProgramu", "PRACA MAGISTERSKA KAMIL RICHTER PB WE 2022"))
        self.Label_t_otoczenia.setText(_translate("menuProgramu", "Temperatura\n"
"otoczenia [C]"))
        self.Label_rodzaj_izolacji.setText(_translate("menuProgramu", "Rodzaj izolacji:"))
        self.rodzaj_izolacji_Input.setItemText(0, _translate("menuProgramu", "Polwinit (PVC)"))
        self.rodzaj_izolacji_Input.setItemText(1, _translate("menuProgramu", "Polietylen usieciowany (XLPE)"))
        #self.rodzaj_izolacji_Input.setItemText(2, _translate("menuProgramu", "Guma etylenowo-propelynowe (EPR)"))
        self.Label_rodzaj_zabezpieczenia.setText(_translate("menuProgramu", "Rodzaj zabezpieczenia:"))
        self.rodzaj_zabezpieczenia_Input.setItemText(0, _translate("menuProgramu", "Wyłącznik instalacyjny"))
        self.rodzaj_zabezpieczenia_Input.setItemText(1, _translate("menuProgramu", "Bezpiecznik"))

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    ekranStartowy = QtWidgets.QMainWindow()
    ui = Ui_ekranStartowy()
    ui.setupUi(ekranStartowy)
    ekranStartowy.show()
    sys.exit(app.exec_())
