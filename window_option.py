import sys
import os
from PyQt5 import QtCore, QtWidgets
from PyQt5.QtGui import QTextCursor, QFontDatabase
from PyQt5.QtCore import Qt, QDateTime, QThread, pyqtSignal
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QGridLayout, QMessageBox, QFileDialog,
              QLabel, QLineEdit, QPushButton, QComboBox, QCheckBox, QDateTimeEdit, QDialog,
              QTextEdit, QDoubleSpinBox, QSpinBox, QTableWidgetItem, QHeaderView)
from docx.enum.style import WD_STYLE_TYPE

import _G, configmanager
from docx import Document
from collections import defaultdict
from util import *

class WindowOption(QDialog):
  def __init__(self, parent=None):
    super().__init__(parent)
    self.parent = parent
    self.setWindowFlag(QtCore.Qt.WindowContextHelpButtonHint, False)
    self.setupUi()
    self.load_options()
  
  def setupUi(self):
    self.setObjectName("Options")
    self.setWindowTitle("選項")
    self.resize(200, 200)
    
    self.main_layout = QGridLayout()
    self.setLayout(self.main_layout)
    
    self.font_size = QSpinBox()
    self.font_chinese = QComboBox()
    self.font_english = QComboBox()
    self.font_risktitle_size = QSpinBox()
    self.style_paragraph = QComboBox()
    self.style_table     = QComboBox()

    self.font_size.setMinimum(_G.FontSizeMinMax[0])
    self.font_size.setMaximum(_G.FontSizeMinMax[1])
    self.font_risktitle_size.setMinimum(_G.FontSizeMinMax[0])
    self.font_risktitle_size.setMaximum(_G.FontSizeMinMax[1])

    self.font_chinese.addItems(_G.Fonts)
    self.font_english.addItems(_G.Fonts)

    self.option_dict = { # Name, Key
      self.font_size: ['字體大小 :', 'NormalFontSize'], 
      self.font_chinese: ['中文字體 :', 'FontChinese'], 
      self.font_english: ['英文字體 :', 'FontEnglish'],
      self.font_risktitle_size: ['風險名稱字體大小 :', 'RiskTitleFontSize'],
      self.style_paragraph: ['段落字體造型', 'ParagStyle'],
      self.style_table: ['表格字體造型', 'TableStyle'],
    }

    _cnt = 0
    for field, val in self.option_dict.items():
      self.main_layout.addWidget(QLabel(val[0]), _cnt, 0, 1, 1, Qt.AlignTop)
      self.main_layout.addWidget(field, _cnt, 1, 1, 1, Qt.AlignTop)
      _cnt += 1

    self.add_confirm_buttons()
    self.connect_buttons()

  def add_confirm_buttons(self):
    self.buttonBox = QtWidgets.QDialogButtonBox(self)
    self.buttonBox.setOrientation(QtCore.Qt.Horizontal)
    self.main_layout.addWidget(self.buttonBox, 0xff, 1, 1, 1, alignment=QtCore.Qt.AlignRight|QtCore.Qt.AlignBottom)
    self.buttonBox.setStandardButtons(QtWidgets.QDialogButtonBox.Cancel|QtWidgets.QDialogButtonBox.Ok)
    self.buttonBox.setObjectName("buttonBox")

  def connect_buttons(self):
    self.buttonBox.accepted.connect(lambda: safe_execute_func(self.on_accept))
    self.buttonBox.rejected.connect(self.on_reject)
    self.buttonBox.button(QtWidgets.QDialogButtonBox.Ok).setText("確認")
    self.buttonBox.button(QtWidgets.QDialogButtonBox.Cancel).setText("取消")
    QtCore.QMetaObject.connectSlotsByName(self)

  def load_styles(self, file):
    self.style_paragraph.clear()
    self.style_table.clear()
    self.style_paragraph.addItem('')
    self.style_table.addItem('')
    try:
      doc = Document(file)
    except Exception:
      msg = f"檔案 {file} 無法讀取，請確認檔案路徑是否正確且應用程式有權限讀取!"
      QMessageBox.warning(self, _G.MsgError, msg, QMessageBox.Ok)
      return False
    st_size,sp_size = 1,1 # empty value
    st_index,sp_index = -1, -1
    
    for s in doc.styles:
      if s.type == WD_STYLE_TYPE.TABLE:
        self.style_table.addItem(s.name)
        if s.name == _G.Config['TableStyle']:
          st_index = st_size
        st_size += 1
      elif s.type == WD_STYLE_TYPE.PARAGRAPH:
        self.style_paragraph.addItem(s.name)
        if s.name == _G.Config['ParagStyle']:
          sp_index = sp_size
        sp_size += 1
    if st_index >= 0:
      self.style_table.setCurrentIndex(st_index)
    if sp_index >= 0:
      self.style_paragraph.setCurrentIndex(sp_index)
    return True

  def on_accept(self):
    new_config = {}
    for field, val in self.option_dict.items():
      ftype = type(field)
      data  = None
      if ftype == QLineEdit:
        data = field.text()
      elif ftype == QSpinBox or ftype == QDoubleSpinBox:
        data = field.value()
      elif ftype == QComboBox:
        data = field.currentText()
      if data:
        new_config[val[1]] = str(data)
    
    configmanager.change_config(_G.ConfigFile, new_config)
    if not _G.load_config():
      QMessageBox.warning(self.parent, _G.MsgError, "選項讀取時發生錯誤，將採用預設值", QMessageBox.Ok)
    self.hide()

  def on_reject(self):
    self.hide()

  def load_options(self):
    _cfg = configmanager.load_all_config(_G.ConfigFile)
    cfg = defaultdict(lambda: '')
    for k, v in _cfg.items():
      if v:
        cfg[k] = str(v)
    
    for field, val in self.option_dict.items():
      key = val[1]
      if key in cfg:
        ftype = type(field)
        if ftype == QLineEdit:
          field.setText(cfg[key])
        elif ftype == QSpinBox or ftype == QDoubleSpinBox:
          field.setValue(configmanager.convert_type(cfg[key]))
        elif ftype == QComboBox:
          _idx = -1
          for i, fname in enumerate(_G.Fonts):
            if fname == cfg[key]:
              _idx = i
              break
          if _idx >= 0:
            field.setCurrentIndex(_idx)