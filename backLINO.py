import math
from math import sqrt
import pandas as pd
from PyQt5 import uic
import sys
from fronLINO import Ui_Dialog
from PyQt5 import QtWidgets
import docx

data = pd.read_excel('Профиль.xlsm')
doc = docx.Document(docx = 'my_written_file.docx')
df = pd.DataFrame(data)
rows_count = data.shape[0]
column_count = data.shape[1]
Form, _ = uic.loadUiType("1.ui")

class Ui(QtWidgets.QDialog, Form):
    def __init__(self):
        super(Ui, self).__init__()
        self.setupUi(self)
        
app = QtWidgets.QApplication(sys.argv)
Dialog = QtWidgets.QDialog()
ui = Ui_Dialog()
ui.setupUi(Dialog)
Dialog.show()

def df_to_list(df):
    lst = []
    for i in range (2, len(df.index)):
        lst.append(df.values[i])
    listoflists = []
    for j in range (0, df.count()[0]):
        proflist_data = []
        for k in range (0, len(data.columns)):
            proflist_data.append(lst[j][k])
        listoflists.append(proflist_data)
    return listoflists
    
def getResult():
    if  len(ui.lineEdit_2.text()) == 0:
        ui.label.setText(f'введите предел текучести')
        ui.label_2.setText(f'')
        ui.label_3.setText(f'')
        return
    else:
        R = float(ui.lineEdit_2.text())
    if  len(ui.lineEdit_3.text()) == 0:
        ui.label.setText(f'введите длину пролета')
        ui.label_2.setText(f'')
        ui.label_3.setText(f'')
        return
    else:
        l = float(ui.lineEdit_3.text())
    if  len(ui.lineEdit_4.text()) == 0:
        ui.label.setText(f'введите количество пролетов')
        ui.label_2.setText(f'')
        ui.label_3.setText(f'')
        return
    else:
        p = float(ui.lineEdit_4.text())
        if p > 3 or p < 1:
            ui.label.setText(f'введите количество пролетов от 1 до 3х')
            ui.label_2.setText(f'')
            ui.label_3.setText(f'')
            return float(ui.lineEdit_4.text())
    if  len(ui.lineEdit_5.text()) == 0:
        ui.label.setText(f'введите ширину опоры')
        ui.label_2.setText(f'')
        ui.label_3.setText(f'')
        return    
    else:    
        b = float(ui.lineEdit_5.text())
    paras = doc.paragraphs
    proflist = df_to_list(df)
    for i in range (0, len(proflist)):
        if ui.comboBox.currentText() == f'{proflist[i][0]}\n':
            row = i + 2
            ix1 = data.iloc[row, 5]
            ix2 = data.iloc[row, 2]
            wx1 = data.iloc[row, 3]
            wx2 = data.iloc[row, 6]
            t = data.iloc[row, 1]
            h = data.iloc[row, 9]
            sd = data.iloc[row, 10]
            sp = data.iloc[row, 11]
            bd = data.iloc[row, 12]
            spf = data.iloc[row, 13]
            emax = data.iloc[row, 14]
            emin = data.iloc[row, 15]
            r = data.iloc[row, 16]
            h0 = data.iloc[row, 17]
            n = data.iloc[row, 18]
            Isum = data.iloc[row, 19]
            name = proflist[i][0]
    if p == 1:
        if ui.comboBox_2.currentText() == 'широкие':
            q1 = round((wx2 * R / (0.125 *l *l * 9.807)), 2)
            q2 = round((wx2 * R / (0.125 *l *l * 9.807)), 2)
        elif ui.comboBox_2.currentText() == 'узкие':
            q1 = round((wx1 * R / (0.125 *l *l * 9.807)), 2)
            q2 = round((wx1 * R / (0.125 *l *l * 9.807)), 2)             
    if p == 2:
        if ui.comboBox_2.currentText() == 'широкие':
            q2 = round((wx2 * R / (0.0625 *l *l * 9.807)), 2)
            q1 = round((wx1 * R / (0.125 *l *l * 9.807)), 2)
        elif ui.comboBox_2.currentText() == 'узкие':
            q2 = round((wx1 * R / (0.0625 *l *l * 9.807)), 2)
            q1 = round((wx2 * R / (0.125 *l *l * 9.807)), 2)
    if p == 3:
        if ui.comboBox_2.currentText() == 'широкие': 
            q2 = round((wx2 * R / (0.075 *l *l * 9.807)), 2)
            q1 = round((wx1 * R / (0.1 *l *l * 9.807)), 2)
        elif ui.comboBox_2.currentText() == 'узкие':
            q2 = round((wx1 * R / (0.075 *l *l * 9.807)), 2)
            q1 = round((wx2 * R / (0.1 *l *l * 9.807)), 2)
    E = 210000
    sina = math.sin(math.pi*data.iloc[row, 20]/180)
    kt = 5.34 + (2.1 / t) * pow((Isum / sd), 1/3)
    koef_zapasa = 1
    if p == 1:
        C0 = 3
        Cr = 0.08
        Cb = 0.7
        Ch = 0.055
    if (p == 2) or (p == 3) :
        C0 = 8
        Cr = 0.1
        Cb = 0.17
        Ch = 0.004
    lyambda1 = 0.346 * (sd / t) * sqrt(5.34 * R / (kt * E))
    lyambda0 = 0.346 * (sp / t) * sqrt((R / E))
    if lyambda0 > lyambda1:
        lyambda = lyambda0
    else:
        lyambda = lyambda1
    if lyambda <= 0.83:
        Rs = 0.58 * R
    elif 0.83 < lyambda < 1.4:
        Rs = 0.48 * R / lyambda
    else:
        Rs = 0.67 * R / pow(lyambda, 2)
    Qw = koef_zapasa * h * t * Rs / (sina * 9.807)
    if p == 1:
        qps = n * Qw / (0.5 * l)
        Q = (0.5 * q2 * l) / n 
    if p == 2:
        qps = n * Qw / (0.625 * l)
        Q = (0.625 * q2 * l) / n 
    if p == 3:
        qps = n * Qw / (0.6 * l)
        Q = (0.6 * q2 * l) / n 
    Qwp = (koef_zapasa * C0 * R * sina * pow(t, 2) * (1 - Cr * sqrt(r / t)) * (1 + Cb * sqrt(b / t)) * (1 - Ch * sqrt(h0 / t))) / 9.807
    kas1 = 1.45 - 0.05 * emax / t
    kas2 = 0.95 + 35000 * pow(t, 2) * emin / (spf * pow(bd, 2))
    if kas1 < kas2:
        kas = kas1
    else:
        kas = kas2
    Qwwp = kas * Qwp
    if p == 1:
        qust = round((n * Qwwp / (0.5 * l)), 2)
    if p == 2:
        qust = round((n * Qwwp / (1.25 * l)), 2)
    if p == 3:
        qust = round((n * Qwwp / (1.1 * l)), 2)
    if l <= 1:
        x = 120
    elif l > 1 and l <= 3:
        x = 120 + (((l - 1) / (3 - 1)) * (150 - 120))
    elif l > 3 and l <= 6:
        x = 150 + (((l - 3) / (6 - 3)) * (200 - 150))
    elif l > 6 and l <= 12:
        x = 200 + (((l - 6) / (12 - 6)) * (250 - 200))
    elif l > 12 and l <= 24:
        x = 250 + (((l - 12) / (24 - 12)) * (300 - 250))
    else:
        x = 300
    if p == 1:
        if ui.comboBox_2.currentText() == 'широкие':
            qj = round((1.4 * (l * ix1 * E * pow(10, -2) * 384 / (x * 5 * pow(l, 4))) / 9.807), 2)
            qjn = round( (l * ix1 * E * pow(10, -2) * 384 / (x * 5 * pow(l, 4)) / 9.807), 2)
        elif ui.comboBox_2.currentText() == 'узкие':
            qj = round((1.4 * (l * ix2 * E * pow(10, -2) * 384 / (x * 5 * pow(l, 4))) / 9.807), 2)
            qjn = round( (l * ix2 * E * pow(10, -2) * 384 / (x * 5 * pow(l, 4)) / 9.807), 2)
    if p == 2:
        if ui.comboBox_2.currentText() == 'широкие':
            qj = round((1.4 *(l * ix1 * E * pow(10, -2) / (x * 0.0052   * pow(l, 4))) / 9.807), 2)
            qjn = round((l * ix1 * E * pow(10, -2) / (x * 0.0052   * pow(l, 4)) / 9.807), 2)
        elif ui.comboBox_2.currentText() == 'узкие':
            qj = round((1.4 *(l * ix2 * E * pow(10, -2) / (x * 0.0052   * pow(l, 4))) / 9.807), 2)
            qjn = round((l * ix2 * E * pow(10, -2) / (x * 0.0052   * pow(l, 4)) / 9.807), 2)
    if p == 3:
        if ui.comboBox_2.currentText() == 'широкие':        
            qj = round((1.4 *(l * ix1 * E * pow(10, -2) / (x * 0.00675 * pow(l, 4))) / 9.807), 2)
            qjn = round((l * ix1 * E * pow(10, -2) / (x * 0.00675 * pow(l, 4)) / 9.807), 2)
        elif ui.comboBox_2.currentText() == 'узкие':
            qj = round((1.4 *(l * ix2 * E * pow(10, -2) / (x * 0.00675 * pow(l, 4))) / 9.807), 2)
            qjn = round((l * ix2 * E * pow(10, -2) / (x * 0.00675 * pow(l, 4)) / 9.807), 2)
    if ui.comboBox_2.currentText() == 'широкие':
        Mr = wx1 * R
    elif ui.comboBox_2.currentText() == 'узкие':
        Mr = wx2 * R
    if p == 1:
        qsovm = round(n * Qwwp/l)
    if p == 2:
        qsovm = round(1.25/(((9.807*0.125*pow(l, 2))/Mr) + ((1.25*l)/(n * Qwwp))))
    if p == 3:
        qsovm = round(1.25/(((9.807*0.1*pow(l, 2))/Mr) + ((1.1*l)/(n * Qwwp))))
    ui.label.setText(f'Прочность: { q1 }кг/м2')
    ui.label_2.setText(f'Жесткость: { qj }кг/м2,       нормативная    { qjn }кг/м2')
    ui.label_3.setText(f'Устойчивость: { qust }кг/м2, совм. действие { qsovm }кг/м2' )
    for para in paras:
        para.text = para.text.replace('list', name)
        para.text = para.text.replace('thickness', str(t))
        para.text = para.text.replace('treshold', str(R))
        para.text = para.text.replace('resist', str(R))
        para.text = para.text.replace('quantaty', str(round(p)))
        para.text = para.text.replace('spanlength', str(l))
        para.text = para.text.replace('width', str(l))
        para.text = para.text.replace('deflection', str(round(x,2)))
        if ui.comboBox_2.currentText() == 'широкие':
            para.text = para.text.replace('squeeze', 'широкие')
        elif ui.comboBox_2.currentText() == 'узкие':
            para.text = para.text.replace('squeeze', 'узкие')
        if ui.comboBox_2.currentText() == 'широкие':
            para.text = para.text.replace('inertia', str(ix2))
        elif ui.comboBox_2.currentText() == 'узкие':
            para.text = para.text.replace('inertia', str(ix1))
        if ui.comboBox_2.currentText() == 'широкие':
            para.text = para.text.replace('moment', str(wx2))
        elif ui.comboBox_2.currentText() == 'узкие':
            para.text = para.text.replace('moment', str(wx1))
        para.text = para.text.replace('strength', str(q1))
        para.text = para.text.replace('rigidity', str(qj))
        para.text = para.text.replace('sustainability', str(qust))
        para.text = para.text.replace('local', str(qsovm))
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = docx.shared.Pt(12)   
    doc.save(f'{name}.docx')
     
ui.pushButton.clicked.connect(getResult)
sys.exit(app.exec_())
