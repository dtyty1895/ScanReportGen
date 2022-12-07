# -*- coding: utf-8 -*-
# version: 2.0.1

# Form implementation generated from reading ui file 'PdfToWord.ui'
#
# Created by: PyQt5 UI code generator 5.14.1
#
# WARNING! All changes made in this file will be lost!

from fileinput import filename
import sys
import time
import os

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtGui import QFontDatabase
from PyQt5.QtWidgets import QWidget, QMainWindow, QApplication
from PyQt5.QtWidgets import QFileDialog, QMessageBox, QCheckBox,QDateTimeEdit,QComboBox
from PyQt5.QtCore import QDateTime
from docx import Document
from pptx import Presentation
import pdfplumber as pdfp
from pandas import read_excel
from webscan_gen import generate_report, silent_generate_report
from threading import Thread
import pickle
from _G import log_debug, log_error, log_warning, log_info
import _G
from window_option import WindowOption
from util import *
import random
from webscan_ppt import *
# Worker async thread for processing documents
worker_thread = None
    
def open_external_file(path):
  try:
    if sys.platform.startswith('linux'):
      os.system(f"xdg-open \"{path}\"")  
    elif sys.platform == 'win32':
      os.system(f"start /b \"\" \"{path}\"")
    else:
      os.system(f"start \"{path}\"")  
    return True
  except Exception:
    return False

class Ui_WebScanGen(object):
  def setupUi(self, WebScanGen):
    WebScanGen.setObjectName("WebScanGen")
    WebScanGen.resize(720, 500)
    WebScanGen.setMinimumSize(QtCore.QSize(720, 500))
    self.load_fonts()
    self.setup_helper_window()
    self.setup_option_window()


    self.setWindowTitle(f"{_G.WindowTitle} {_G.Version}")
    self.verticalLayout = QtWidgets.QVBoxLayout(WebScanGen)
    self.verticalLayout.setObjectName("verticalLayout")
    
    self.horizontalLayout_5 = QtWidgets.QHBoxLayout()
    self.horizontalLayout_5.setObjectName("horizontalLayout_5")
    
    self.label = QtWidgets.QLabel(WebScanGen)
    self.label.setObjectName("label")
    self.horizontalLayout_5.addWidget(self.label)
    self.Dev_in = QtWidgets.QLineEdit(WebScanGen)
    self.Dev_in.setObjectName("Dev_in")
    self.horizontalLayout_5.addWidget(self.Dev_in)
    self.Dev_btn = QtWidgets.QPushButton(WebScanGen)
    self.Dev_btn.setObjectName("Dev_btn")
    self.horizontalLayout_5.addWidget(self.Dev_btn)
    self.verticalLayout.addLayout(self.horizontalLayout_5)
    
    self.horizontalLayout_7 = QtWidgets.QHBoxLayout()
    self.horizontalLayout_7.setObjectName("horizontalLayout_7")
    self.label_2 = QtWidgets.QLabel(WebScanGen)
    self.label_2.setObjectName("label_2")
    self.horizontalLayout_7.addWidget(self.label_2)
    self.Owa_in = QtWidgets.QLineEdit(WebScanGen)
    self.Owa_in.setObjectName("Owa_in")
    self.horizontalLayout_7.addWidget(self.Owa_in)
    self.Owa_btn = QtWidgets.QPushButton(WebScanGen)
    self.Owa_btn.setObjectName("Owa_btn")
    self.horizontalLayout_7.addWidget(self.Owa_btn)
    self.verticalLayout.addLayout(self.horizontalLayout_7)
    self.horizontalLayout = QtWidgets.QHBoxLayout()
    self.horizontalLayout.setObjectName("horizontalLayout")
    self.Excel = QtWidgets.QLabel(WebScanGen)
    self.Excel.setObjectName("Excel")
    self.horizontalLayout.addWidget(self.Excel)
    self.Excel_in = QtWidgets.QLineEdit(WebScanGen)
    self.Excel_in.setObjectName("Excel_in")
    self.horizontalLayout.addWidget(self.Excel_in)
    self.Excel_btn = QtWidgets.QPushButton(WebScanGen)
    self.Excel_btn.setMinimumSize(QtCore.QSize(100, 0))
    self.Excel_btn.setObjectName("Excel_btn")
    self.horizontalLayout.addWidget(self.Excel_btn)
    self.verticalLayout.addLayout(self.horizontalLayout)
    
    self.horizontalLayout_2 = QtWidgets.QHBoxLayout()
    self.horizontalLayout_2.setObjectName("horizontalLayout_2")
    self.Word = QtWidgets.QLabel(WebScanGen)
    self.Word.setObjectName("Word")
    self.horizontalLayout_2.addWidget(self.Word)
    self.Word_in = QtWidgets.QLineEdit(WebScanGen)
    self.Word_in.setObjectName("Word_in")
    self.horizontalLayout_2.addWidget(self.Word_in)
    self.Word_btn = QtWidgets.QPushButton(WebScanGen)
    self.Word_btn.setMinimumSize(QtCore.QSize(100, 0))
    self.Word_btn.setObjectName("Word_btn")
    self.horizontalLayout_2.addWidget(self.Word_btn)
    self.verticalLayout.addLayout(self.horizontalLayout_2)

    self.horizontalLayout_8 = QtWidgets.QHBoxLayout()
    self.horizontalLayout_8.setObjectName("horizontalLayout_8")
    self.ppt = QtWidgets.QLabel(WebScanGen)
    self.ppt.setObjectName("ppt")
    self.horizontalLayout_8.addWidget(self.ppt)
    self.ppt_in = QtWidgets.QLineEdit(WebScanGen)
    self.ppt_in.setObjectName("ppt_in")
    self.horizontalLayout_8.addWidget(self.ppt_in)
    self.ppt_btn = QtWidgets.QPushButton(WebScanGen)
    self.ppt_btn.setMinimumSize(QtCore.QSize(100, 0))
    self.ppt_btn.setObjectName("ppt_btn")
    self.horizontalLayout_8.addWidget(self.ppt_btn)
    self.verticalLayout.addLayout(self.horizontalLayout_8)
    
    
    self.horizontalLayout_9 = QtWidgets.QHBoxLayout()
    self.horizontalLayout_9.setObjectName("horizontalLayout_9")
    self.pdf_dir = QtWidgets.QLabel(WebScanGen)
    self.pdf_dir.setObjectName("pdf_dir")
    self.horizontalLayout_9.addWidget(self.pdf_dir)
    self.pdf_dir_in = QtWidgets.QLineEdit(WebScanGen)
    self.pdf_dir_in.setObjectName("pdf_dir_in")
    self.horizontalLayout_9.addWidget(self.pdf_dir_in)
    self.pdf_dir_btn = QtWidgets.QPushButton(WebScanGen)
    self.pdf_dir_btn.setMinimumSize(QtCore.QSize(100, 0))
    self.pdf_dir_btn.setObjectName("pdf_dir_btn")
    self.horizontalLayout_9.addWidget(self.pdf_dir_btn)
    self.verticalLayout.addLayout(self.horizontalLayout_9)
    
    
    
    self.horizontalLayout_3 = QtWidgets.QHBoxLayout()
    self.horizontalLayout_3.setObjectName("horizontalLayout_3")
    self.CompanyName = QtWidgets.QLabel(WebScanGen)
    self.CompanyName.setObjectName("CompanyName")
    self.horizontalLayout_3.addWidget(self.CompanyName)
    self.CompanyName_in = QtWidgets.QLineEdit(WebScanGen)
    self.CompanyName_in.setObjectName("CompanyName_in")
    self.horizontalLayout_3.addWidget(self.CompanyName_in)
    self.CompanyNameAbbr = QtWidgets.QLabel(WebScanGen)
    self.CompanyNameAbbr.setObjectName("CompanyNameAbbr")
    self.horizontalLayout_3.addWidget(self.CompanyNameAbbr)
    self.CompanyNameAbbr_in = QtWidgets.QLineEdit(WebScanGen)
    self.CompanyNameAbbr_in.setObjectName("CompanyNameAbbr_in")
    self.horizontalLayout_3.addWidget(self.CompanyNameAbbr_in)
    
    self.scheduler = QtWidgets.QLabel(WebScanGen)
    self.scheduler.setObjectName("Date")
    self.horizontalLayout_3.addWidget(self.scheduler)
    # self.in_scheduler = QtWidgets.QLineEdit(WebScanGen)
    self.in_scheduler = QDateTimeEdit(QDateTime.currentDateTime())
    self.in_scheduler.setCalendarPopup(True)
    self.in_scheduler.setDisplayFormat('yyyy-MM-dd')
    self.in_scheduler.setObjectName("in_scheduler")
    self.horizontalLayout_3.addWidget(self.in_scheduler)
    
    self.verticalLayout.addLayout(self.horizontalLayout_3)


    self.textEdit = QtWidgets.QTextEdit(WebScanGen)
    self.textEdit.setObjectName("textEdit")
    self.verticalLayout.addWidget(self.textEdit)
    self.horizontalLayout_4 = QtWidgets.QHBoxLayout()
    self.horizontalLayout_4.setObjectName("horizontalLayout_4")
    self.Save = QtWidgets.QLabel(WebScanGen)
    self.Save.setObjectName("Save")
    self.horizontalLayout_4.addWidget(self.Save)
    self.Save_in = QtWidgets.QLineEdit(WebScanGen)
    self.Save_in.setObjectName("Save_in")
    self.horizontalLayout_4.addWidget(self.Save_in)
    # self.Savefile = QtWidgets.QLabel(WebScanGen)
    # self.Savefile.setObjectName("Savefile")
    # self.horizontalLayout_4.addWidget(self.Savefile)
    # self.Savefile_in = QtWidgets.QLineEdit(WebScanGen)
    # self.Savefile_in.setMaximumSize(QtCore.QSize(100, 16777215))
    # self.Savefile_in.setObjectName("Savefile_in")
    # self.horizontalLayout_4.addWidget(self.Savefile_in)
    self.Save_btn = QtWidgets.QPushButton(WebScanGen)
    self.Save_btn.setMinimumSize(QtCore.QSize(100, 0))
    self.Save_btn.setObjectName("Save_btn")
    self.horizontalLayout_4.addWidget(self.Save_btn)
    self.verticalLayout.addLayout(self.horizontalLayout_4)
    self.horizontalLayout_6 = QtWidgets.QHBoxLayout()
    self.horizontalLayout_6.setObjectName("horizontalLayout_6")
    spacerItem = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
    self.horizontalLayout_6.addItem(spacerItem)
    # self.ch_radio = QtWidgets.QRadioButton(WebScanGen)
    # self.ch_radio.setObjectName("ch_radio")
    # self.ch_radio.setChecked(True)
    # self.horizontalLayout_6.addWidget(self.ch_radio)
    # self.en_radio = QtWidgets.QRadioButton(WebScanGen)
    # self.en_radio.setObjectName("en_radio")
    # self.horizontalLayout_6.addWidget(self.en_radio)
    
    self.Chs_typ = QtWidgets.QLabel(WebScanGen)
    self.Chs_typ.setObjectName("Choose_Type")
    self.horizontalLayout_6.addWidget(self.Chs_typ)
    self.chs_typ_in = QComboBox()
    self.chs_typ_in.currentIndexChanged.connect(self.on_type_select)
    self.chs_typ_in.addItems(['PPT', 'WORD', 'ALL'])
    self.horizontalLayout_6.addWidget(self.chs_typ_in)
    
    self.chk_openfinished = QCheckBox("生成報告後自動開啟")
    self.chk_openfinished.stateChanged.connect(self.on_auto_open)
    self.horizontalLayout_6.addWidget(self.chk_openfinished)
    
    self.btn_info = QtWidgets.QPushButton('使用說明')
    self.btn_info.clicked.connect(self.show_helper)
    self.btn_info.setObjectName("Info_btn")
    self.horizontalLayout_6.addWidget(self.btn_info)

    self.btn_option = QtWidgets.QPushButton('選項')
    self.btn_option.clicked.connect(self.show_option)
    self.btn_option.setObjectName("Option_btn")
    self.horizontalLayout_6.addWidget(self.btn_option)

    self.Clear_btn = QtWidgets.QPushButton(WebScanGen)
    self.Clear_btn.setMinimumSize(QtCore.QSize(100, 0))
    self.Clear_btn.setMaximumSize(QtCore.QSize(100, 16777215))
    self.Clear_btn.setObjectName("Clear_btn")
    self.horizontalLayout_6.addWidget(self.Clear_btn)
    
    self.Excute_btn = QtWidgets.QPushButton(WebScanGen)
    self.Excute_btn.setMinimumSize(QtCore.QSize(100, 0))
    self.Excute_btn.setObjectName("Excute_btn")
    self.horizontalLayout_6.addWidget(self.Excute_btn)
    
    self.verticalLayout.addLayout(self.horizontalLayout_6)
    
    
    

    self.Dev_in.setReadOnly(True)
    self.Owa_in.setReadOnly(True)
    self.Excel_in.setReadOnly(True)
    self.Word_in.setReadOnly(True)
    self.ppt_in.setReadOnly(True)
    self.pdf_dir_in.setReadOnly(True)
    self.retranslateUi(WebScanGen)
    
    self.Dev_btn.clicked.connect(lambda: self.open_file(0))
    self.Owa_btn.clicked.connect(lambda: self.open_file(1))
    self.Excel_btn.clicked.connect(lambda: self.open_file(2))
    self.Word_btn.clicked.connect(lambda: self.open_file(3))
    self.ppt_btn.clicked.connect(lambda: self.open_file(4))
    self.pdf_dir_btn.clicked.connect(lambda: self.open_file(5))
    
    self.Save_btn.clicked.connect(self.save_path)
    self.Excute_btn.clicked.connect(self.input_data)

    self.Clear_btn.clicked.connect(self.clear_input)

    QtCore.QMetaObject.connectSlotsByName(WebScanGen)
    
    self.auto_open = False
    self.textEdit.setReadOnly(True)
    self.active_buttons = [
      self.Dev_btn, self.Owa_btn, self.Excel_btn, self.Word_btn, self.Excute_btn, 
      self.Clear_btn, self.Save_btn, self.Dev_in, self.Owa_in, self.Excel_in,
      self.Word_in,self.ppt_in, self.Save_in, self.CompanyName_in, self.btn_option,
      self.CompanyNameAbbr_in, self.ppt_in, self.ppt_btn, self.pdf_dir_in, self.pdf_dir_btn,
      self.Chs_typ, self.chs_typ_in
    ]
    self.load_cache()
    fname = self.Word_in.text().strip()
    if self.check_file_unreadable(fname):
      self.Word_in.clear()
    elif fname:
      self.window_option.load_styles(fname)
    
  def retranslateUi(self, WebScanGen):
    _translate = QtCore.QCoreApplication.translate
    WebScanGen.setWindowTitle(_translate("WebScanGen", "Excel To Word"))
    self.label.setText(_translate("WebScanGen", "Developer PDF"))
    self.Dev_btn.setText(_translate("WebScanGen", "選取 PDF"))
    self.label_2.setText(_translate("WebScanGen", "   OWASP PDF"))
    self.Owa_btn.setText(_translate("WebScanGen", "選取 PDF"))
    self.Excel.setText(_translate("WebScanGen", "      翻譯 Excel"))
    self.Excel_btn.setText(_translate("WebScanGen", "選取 Excel"))
    self.Word.setText(_translate("WebScanGen", "      模板 Word"))
    self.Word_btn.setText(_translate("WebScanGen", "選取 Word"))
    self.ppt.setText(_translate("WebScanGen", "      模板 PPT"))
    self.ppt_btn.setText(_translate("WebScanGen", "選取 PPT"))
    self.pdf_dir.setText(_translate('WebScanGen', 'PDF資料夾'))
    self.pdf_dir_btn.setText(_translate("WebScanGen" , '選取資料夾'))
    self.Save.setText(_translate("WebScanGen", "儲存至："))
    # self.Savefile.setText(_translate("WebScanGen", "檔案名"))
    self.Save_btn.setText(_translate("WebScanGen", "儲存路徑"))
    self.Clear_btn.setText(_translate("WebScanGen", "清除"))
    self.Excute_btn.setText(_translate("WebScanGen", "執行"))
    self.CompanyName.setText(_translate("WebScanGen", "公司名稱"))
    self.CompanyNameAbbr.setText(_translate("WebScanGen", "公司簡稱"))
    self.scheduler.setText(_translate("Date","日期"))
    self.Chs_typ.setText(_translate("Choose_Type","類型"))
    
  def on_type_select(self, index):
    global word_Active , ppt_Active    
    
    if index == 0:
      ppt_Active=True
      word_Active=False
      self.label.setEnabled(False)
      self.Dev_btn.setEnabled(False)
      self.Owa_btn.setEnabled(False)
      self.label_2.setEnabled(False)
      self.Word.setEnabled(False)
      self.Word_btn.setEnabled(False)
      self.ppt.setEnabled(True)
      self.ppt_btn.setEnabled(True)
      self.pdf_dir.setEnabled(True)
      self.pdf_dir_btn.setEnabled(True)
    elif index == 1:
      ppt_Active=False
      word_Active=True
      self.label.setEnabled(True)
      self.Dev_btn.setEnabled(True)
      self.Owa_btn.setEnabled(True)
      self.label_2.setEnabled(True)
      self.ppt.setEnabled(False)
      self.ppt_btn.setEnabled(False)
      self.Word.setEnabled(True)
      self.Word_btn.setEnabled(True)
      self.pdf_dir.setEnabled(False)
      self.pdf_dir_btn.setEnabled(False)
    elif index == 2 :
      
      ppt_Active=True
      word_Active=True
      self.label.setEnabled(True)
      self.Dev_btn.setEnabled(True)
      self.Owa_btn.setEnabled(True)
      self.label_2.setEnabled(True)
      self.ppt.setEnabled(True)
      self.ppt_btn.setEnabled(True)
      self.Word.setEnabled(True)
      self.Word_btn.setEnabled(True)
      self.pdf_dir.setEnabled(True)
      self.pdf_dir_btn.setEnabled(True)
  
  def load_fonts(self):
    for font in QFontDatabase().families():
      _G.Fonts.append(font)

  def setup_helper_window(self):
    self.window_helper = QMessageBox()
    self.window_helper.setStyleSheet("QLabel{min-width: 600px; font-size: 14px}")
    self.window_helper.setWindowTitle("使用說明")
    self.window_helper.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
    self.window_helper.setText('''
    <p>程式版本: $TAG_SOFTWARE_VERSION$<p>
    <p>此報告生成器適用於 Acunetix v13/v14 所產出之 Developer 及 OWASP PDF，並須自行準備公版(模板)以自動產出報告。
    若需要自動產出部分的幫助，請閱讀外部完ParagStyle整使用說明文件。</p>
    <hr>
    <p>UI 說明：</p>
    <table border=0 cellpadding=0 cellspacing=10>
    <tr><td>Developer PDF：按下欄位右方的瀏覽可選擇檔案，應選擇掃描所產出的 Developer PDF。</td></tr>
    <tr><td>OWASP PDF：功能同上；應選擇掃描所產出的 OWASP PDF。</td></tr>
    <tr><td>翻譯 Excel：事先準備好的弱點對應翻譯表，格式詳見外部說明文件；若無翻譯資料將填入原文。</td></tr>
    <tr><td>模板 Word：事先準備好的模板，程式將會依據插入標籤自動填入 PDF 資料。</td></tr>
    <tr><td>公司名稱/簡稱：掃描的單位或公司名稱/簡稱。</td></tr>
    <tr><td>儲存路徑：功能同 "另存新檔"，點選後將選擇產出報告檔案的儲存目的地及名稱。</td></tr>
    <tr><td>清除：清除所有輸入欄位</td></tr>
    <tr><td>執行：程式將會開始執行，部分資訊將會顯示在按鍵上方的資訊欄位。</td></tr>
    </table>
    <hr>
    <p>程式執行期間輸入欄位將會鎖定，不過仍可以選擇是否執行完成後自動開啟檔案。</p>
    '''.replace('$TAG_SOFTWARE_VERSION$', _G.Version))
    btn_ok = self.window_helper.button(QMessageBox.Yes)
    btn_ok.setText(" 打開外部使用說明文件 ")
    btn_ok.clicked.connect(self.open_helper_doc)
    btn_no = self.window_helper.button(QMessageBox.No)
    btn_no.setText("關閉")
    self.window_helper.hide()

  def show_helper(self):
    self.window_helper.show()

  def setup_option_window(self):
    self.window_option = WindowOption(self)
    self.window_option.setModal(True)
    self.window_option.hide()
  
  def show_option(self):
    self.window_option.show()
    self.window_option.load_options()

  def open_helper_doc(self):
    _ok = open_external_file(_G.READMEFile)
    if not _ok:
      QMessageBox.warning(self, _G.MsgError, f"無法開啟說明文件, 請自行閱讀線上文件 '{_G.READMEFile}'", QMessageBox.Ok)

  def open_file(self, field_id):
    log_debug(f"File selection field id: {field_id}")
    file_fields = [self.Dev_in, self.Owa_in, self.Excel_in, self.Word_in, self.ppt_in, self.pdf_dir_in]
    guess_fid = -1

    # Open Excel
    if field_id == 2:
      file_name, _ = QFileDialog.getOpenFileName(self, "選取檔案", "./", _G.FileTypeExcel)
      print('file_name : '+file_name)
      if not file_name:
        return
      if is_file_corrupted(file_name, _G.FileTypeExcel):
        return self.on_file_corrupted(file_name)
    # Open Word
    elif field_id == 3:
      file_name, _ = QFileDialog.getOpenFileName(self, "選取檔案", "./", _G.FileTypeWord)
      if not file_name:
        return
      if is_file_corrupted(file_name, _G.FileTypeWord):
        return self.on_file_corrupted(file_name)
      # self.window_option.load_styles(file_name)
    # Open PDF
    elif field_id == 0 or field_id == 1:
      file_name, _ = QFileDialog.getOpenFileName(self, "選取檔案", "./", _G.FileTypePDF)
      if not file_name:
        return
      if is_file_corrupted(file_name, _G.FileTypePDF):
        return self.on_file_corrupted(file_name)
      elif 'Developer' in file_name:
        guess_fid = 0
      elif 'OWASP' in file_name:
        guess_fid = 1 
        
    elif field_id == 4:
      file_name, _ = QFileDialog.getOpenFileName(self, "選取檔案", "./", _G.FileTypePPT)
      print("filename"+file_name)
      if not file_name:
        return
      if is_file_corrupted(file_name, _G.FileTypePPT):
        return self.on_file_corrupted(file_name)
    elif field_id == 5:
        file_name = QFileDialog.getExistingDirectory(self, caption='選取資料夾', directory=os.getcwd()  )
        _G.pdf_path = file_name
        if not file_name:
            return
          
    if guess_fid >= 0 and guess_fid != field_id:
      qm = QMessageBox
      reply = qm.question(self, _G.MsgInfo, self.get_hint_message(field_id, guess_fid), qm.Yes | qm.No)
      if reply == qm.Yes:
        field_id = guess_fid
    file_fields[field_id].setText(file_name)
    self.dump_cache()

  def on_file_corrupted(self, filename):
    QMessageBox.warning(self, _G.MsgError, '選取的檔案毀損, 無法讀取!', QMessageBox.Ok)
    self.textEdit.append(f"檔案 {filename} 毀損, 無法讀取 :(")

  def save_path(self):
    if word_Active == False:
      dest, _ = QFileDialog.getSaveFileName(self, '另存為...', './', 'PPT Files (*.pptx)')
      self.Save_in.setText(dest)
    elif ppt_Active == False:
      dest, _ = QFileDialog.getSaveFileName(self, '另存為...', './', 'Word Files (*.docx)')
      self.Save_in.setText(dest)
    elif ppt_Active ==True and word_Active == True:
      dest, _ = QFileDialog.getSaveFileName(self, '另存為...', './', 'Word Files (*.docx)')
      dest2, _ = QFileDialog.getSaveFileName(self, '另存為...', './', 'PPT Files (*.pptx)')   
      self.Save_in.setText(dest+'&&'+dest2)
  
  def clear_input(self):
    fields = [self.Dev_in, self.Owa_in, self.Excel_in, self.Word_in, self.ppt_in, self.textEdit, 
    self.CompanyName_in, self.CompanyNameAbbr_in, self.in_scheduler, self.Save_in]
    for field in fields:
      field.clear()

  def input_data(self):
    
    dev_pdf = self.Dev_in.text().strip()
    owa_pdf = self.Owa_in.text().strip()
    excel = self.Excel_in.text().strip()
    word = self.Word_in.text().strip()
    ppt = self.ppt_in.text().strip()
    pdf_dir = self.pdf_dir_in.text().strip()
    company_name = self.CompanyName_in.text().strip()
    company_abbr = self.CompanyNameAbbr_in.text().strip()
    
    if  ppt_Active==True and word_Active==True:
      save = self.Save_in.text().strip().split('&&')
    else :
      save = self.Save_in.text().strip()
    style_table = self.window_option.style_table.currentText()
    style_parag = self.window_option.style_paragraph.currentText()
    date = self.in_scheduler.text().split('-')
    
    print('date ' + date[0])
    _G.ppt_params['YYYY']=date[0]
    _G.ppt_params['MM']=date[1]
    _G.ppt_params['DD']=date[2]
    _G.ppt_params['OOOO']=company_name
    _G.slide_add = 0
    date = date[0]+'/'+date[1]+'/'+date[2]
    log_info("Save location: ", save)

    if word_Active==True and (not dev_pdf or not owa_pdf or not excel or not word or not save):
      QMessageBox.warning(self, "缺少資料", "請確認必要資料是否填入", QMessageBox.Ok)
      return
    elif not company_name or not company_abbr:
      QMessageBox.warning(self, "缺少資料", "請填入公司名稱及簡稱", QMessageBox.Ok)
      return
    # elif not style_table or not style_parag:
    #   QMessageBox.warning(self, "缺少資料", "請選擇 '選項' 中的段落與表格造型", QMessageBox.Ok)
    #   return
    elif  ppt_Active==True and (not pdf_dir or not ppt):
      QMessageBox.warning(self, "缺少資料", "請確認必要資料是否填入", QMessageBox.Ok)
      return
    else:
      self.textEdit.append("Developer PDF : %s" % dev_pdf)
      self.textEdit.append("OWASP PDF : %s" % owa_pdf)            
      self.textEdit.append("翻譯 Excel : %s" % excel)
      self.textEdit.append("模板 Word : %s" % word)
      self.textEdit.append("模板 ppt : %s" % ppt)
      self.textEdit.append("pdf資料夾 : %s" % pdf_dir)
      if ppt_Active==True and word_Active==True:
        self.textEdit.append("Save file Path : %s" % save[0])
        self.textEdit.append("Save file Path : %s" % save[1])
      else:
        self.textEdit.append("Save file Path : %s" % save)
      if ppt_Active==True and word_Active==True: 
        print("bbb")
        # failed_file = self.check_file_unreadable(dev_pdf, owa_pdf, excel, word, save[0],save[1])
        if os.path.isfile(save[0]) or os.path.isfile(save[1]):
          qm = QMessageBox
          reply = qm.question(self, _G.MsgInfo, "存檔位置已有同名檔案存在, 是否覆蓋？", qm.Yes | qm.No)
          if reply == qm.No:
            return
        elif not is_path_writable(save[0]) or not is_path_writable(save[1]):
          QMessageBox.warning(self, _G.MsgError, _G.MsgUnwritable, QMessageBox.Ok)
          return
      else:
        print('aaa')
        # failed_file = self.check_file_unreadable(dev_pdf, owa_pdf, excel, word, save)
        if os.path.isfile(save):
          qm = QMessageBox
          reply = qm.question(self, _G.MsgInfo, "存檔位置已有同名檔案存在, 是否覆蓋？", qm.Yes | qm.No)
          if reply == qm.No:
            return
        elif not is_path_writable(save) :
          QMessageBox.warning(self, _G.MsgError, _G.MsgUnwritable, QMessageBox.Ok)
          return
      # print(failed_file)
      if ppt_Active==True and word_Active==True: 
        failed_file = self.check_file_unreadable(dev_pdf, owa_pdf, excel, word, save[0],save[1])
      else:
        failed_file = self.check_file_unreadable(dev_pdf, owa_pdf, excel, word, save)
      if failed_file:
        msg = f"檔案 {failed_file} 無法讀取，請確認檔案路徑是否正確且應用程式有權限讀取!"
        QMessageBox.warning(self, _G.MsgError, msg, QMessageBox.Ok)
        return
  
      self.disable_buttons()
      QApplication.processEvents()
      
      _G.DocTableStyle = style_table
      _G.DocParagStyle = style_parag

      self.textEdit.append("執行中...\n")
      self.textEdit.moveCursor(QtGui.QTextCursor.End)

      if word_Active == False:
        ppt_worker_thread = Thread(target=silent_generate_ppt_report,
                                   args = [save,ppt,excel],
                                    daemon=True)
        ppt_worker_thread.start()
        has_error = False
        
        while ppt_worker_thread.is_alive():
          QApplication.processEvents()
          if _G.PipeMessages:
            messages = _G.pop_pipe_messages()
            for msg in messages:
              self.textEdit.append(msg)
          if _G.PipeError:
            has_error = True
            err, errinfo = _G.PipeError.popleft()
            handle_exception(err, errinfo)
            self.textEdit.append("---------------\n")
            break
          
          if _G.PipeWarning:
            msg = _G.PipeWarning.popleft()
            if msg == _G.MsgPipeWarnTargetOpened:
              QMessageBox.warning(self, _G.MsgInfo, msg, QMessageBox.Ok)
              _G.PipeSubInfo.append(_G.MsgPipeContinue)
        
      elif ppt_Active == False:
        worker_thread = Thread(
          target=silent_generate_report, 
          args=[dev_pdf, owa_pdf, word, excel, save], 
          kwargs={'company_name': company_name, 'company_abbr': company_abbr, 'date':date},
          daemon=True
        )

        worker_thread.start()
        
        has_error = False
        while worker_thread.is_alive():
          QApplication.processEvents()
          if _G.PipeMessages:
            messages = _G.pop_pipe_messages()
            for msg in messages:
              self.textEdit.append(msg)
          if _G.PipeError:
            has_error = True
            err, errinfo = _G.PipeError.popleft()
            handle_exception(err, errinfo)
            self.textEdit.append("---------------\n")
            break
          
          if _G.PipeWarning:
            msg = _G.PipeWarning.popleft()
            if msg == _G.MsgPipeWarnTargetOpened:
              QMessageBox.warning(self, _G.MsgInfo, msg, QMessageBox.Ok)
              _G.PipeSubInfo.append(_G.MsgPipeContinue)
      elif ppt_Active==True and word_Active==True:
        worker_thread = Thread(
          target=silent_generate_report, 
          args=[dev_pdf, owa_pdf, word, excel, save[0]], 
          kwargs={'company_name': company_name, 'company_abbr': company_abbr},
          daemon=True
        )

        worker_thread.start()
        
        has_error = False
        while worker_thread.is_alive():
          QApplication.processEvents()
          if _G.PipeMessages:
            messages = _G.pop_pipe_messages()
            for msg in messages:
              self.textEdit.append(msg)
          if _G.PipeError:
            has_error = True
            err, errinfo = _G.PipeError.popleft()
            handle_exception(err, errinfo)
            self.textEdit.append("---------------\n")
            break
          
          if _G.PipeWarning:
            msg = _G.PipeWarning.popleft()
            if msg == _G.MsgPipeWarnTargetOpened:
              QMessageBox.warning(self, _G.MsgInfo, msg, QMessageBox.Ok)
              _G.PipeSubInfo.append(_G.MsgPipeContinue)
              
        if has_error:
            QMessageBox.warning(self, _G.MsgError, '程式發生錯誤，詳情請見資訊窗格', QMessageBox.Ok)
        else:
          self.textEdit.append("WORD執行完畢, 檔案已儲存至 "+save[0])
     
        
        ppt_worker_thread = Thread(target=silent_generate_ppt_report,
                                   args = [save[1],ppt,excel],
                                    daemon=True)
        ppt_worker_thread.start()
        has_error = False
        
        while ppt_worker_thread.is_alive():
          QApplication.processEvents()
          if _G.PipeMessages:
            messages = _G.pop_pipe_messages()
            for msg in messages:
              self.textEdit.append(msg)
          if _G.PipeError:
            has_error = True
            err, errinfo = _G.PipeError.popleft()
            handle_exception(err, errinfo)
            self.textEdit.append("---------------\n")
            break
          
          if _G.PipeWarning:
            msg = _G.PipeWarning.popleft()
            if msg == _G.MsgPipeWarnTargetOpened:
              QMessageBox.warning(self, _G.MsgInfo, msg, QMessageBox.Ok)
              _G.PipeSubInfo.append(_G.MsgPipeContinue)
              
        
      if has_error:
        QMessageBox.warning(self, _G.MsgError, '程式發生錯誤，詳情請見資訊窗格', QMessageBox.Ok)
      else:
        self.textEdit.append("執行完畢, 檔案已儲存至 "+save[1])
        self.textEdit.append("請手動更改:\n目錄頁碼")
        QMessageBox.information(self, _G.MsgInfo, '執行完畢！', QMessageBox.Ok)
      self.textEdit.moveCursor(QtGui.QTextCursor.End)
      QApplication.processEvents()
      self.enable_buttons()

      if self.auto_open:
        try:
          if sys.platform.startswith('linux'):
            os.system(f"xdg-open \"{save}\"")  
          elif sys.platform == 'win32':
            os.system(f"start /b \"\" \"{save}\"")
          else:
            os.system(f"start \"{save}\"")  
        except Exception:
          self.textEdit.append("自動開啟不支援當前版本的作業系統, 請手動開啟")
  
  def on_auto_open(self, signal):
    self.auto_open = True if signal > 0 else False

  def disable_buttons(self):
    for btn in self.active_buttons:
      btn.setEnabled(False)

  def enable_buttons(self):
    for btn in self.active_buttons:
      btn.setDisabled(False)
  
  def get_hint_message(self, fid, gid):
    ret = "您現在選擇的檔案欄位是: "
    ret += _G.UIInputFieldNames[fid]
    ret += ",\n但是選擇的檔案似乎是: "
    ret += _G.UIInputFieldNames[gid]
    ret += "\n請問是否要將檔案改成輸入到建議的欄位(而非原本選擇的欄位)？"
    return ret
  
  def check_file_unreadable(self, *files):
    for file in files:
      try:
        with open(file, 'r') as _:
          pass
      except Exception:
        return file
    return False
  
  def load_cache(self):
    if not os.path.isfile(_G.CacheFile):
      return
    try:
      with open(_G.CacheFile, 'rb') as fp:
        dat = pickle.load(fp)
        self.Excel_in.setText(dat['tmp_xls'])
        self.Word_in.setText(dat['tmp_word'])
    except Exception:
      return

  def dump_cache(self):
    dat_old = {'tmp_xls': None, 'tmp_word': None}
    try:
      if os.path.isfile(_G.CacheFile):
        with open(_G.CacheFile, 'rb') as fp:
          dat_old = pickle.load(fp)
    except Exception:
      pass

    try:  
      with open(_G.CacheFile, 'wb') as fp:
        dat = {
          'tmp_xls': self.Excel_in.text(),
          'tmp_word': self.Word_in.text()
        }
        if not dat['tmp_xls'] and dat_old['tmp_xls']:
          dat['tmp_xls'] = dat_old['tmp_xls']
        if not dat['tmp_word'] and dat_old['tmp_word']:
          dat['tmp_word'] = dat_old['tmp_word']
        pickle.dump(dat, fp)
    except Exception:
      return

class MainWindow(QMainWindow, Ui_WebScanGen):
  def __init__(self, parent=None):
    super(MainWindow, self).__init__(parent)
    central_widget = QWidget()
    self.setCentralWidget(central_widget) # new central widget    
    self.setupUi(central_widget)

def is_path_writable(path):
  try:
    with open(path, 'w') as fp:
      fp.write("\0")
    return True
  except Exception:
    return False
  

def is_file_corrupted(file, ftype):
  try:
    if ftype == _G.FileTypePDF:
      pdfp.open(file)
    elif ftype == _G.FileTypeExcel:
      read_excel(file)
    elif ftype == _G.FileTypeWord:
      Document(file)
    elif ftype == _G.FileTypePPT:
      Presentation(file)
  except Exception:
    return True
  return False

if __name__ == "__main__":
  app = QApplication(sys.argv)
  window = MainWindow()
  _G.MainWindow = window
  log_info("Window initialized")
  window.show()
  sys.exit(app.exec_())
    