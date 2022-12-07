from PyQt5.QtWidgets import QMessageBox
from difflib import SequenceMatcher
import traceback
import _G
import collections

def safe_execute_func(func, args=[], kwargs={}):
  try:
    return func(*args, **kwargs)
  except Exception as err:
    err_info = traceback.format_exc()
    handle_exception(err, err_info, _G.MainWindow)
    _G.log_error(f"An error occurred!\n{err_info}")
  return _G.MsgPipeError

def handle_exception(err, errinfo, window=None):
  dmp_ok = dump_errorlog(err, errinfo)
  if window:
    if dmp_ok:
      window.edit_log.append(f"運行過程中產生錯誤, 請將紀錄檔 {dmp_ok} 寄送給開發人員以排除錯誤")
    else:
      window.edit_log.append(f"運行過程中產生錯誤, 請將以下內容複製並寄送給開發人員以排除錯誤")
      window.edit_log.append(str(err))
      window.edit_log.append(errinfo)
    window.edit_log.append("同時請確認您的檔案及內容是正確的.")
    window.edit_log.append(_G.get_witty_comment())
  else:
    if dmp_ok:
      QMessageBox.critical(None, _G.MsgError, f'運行過程中產生錯誤, 請將紀錄檔 {dmp_ok} 寄送給開發人員以排除錯誤, 同時請確認您的檔案及內容是正確的', QMessageBox.Ok)
    else:
      _msg  = '運行過程中產生錯誤, 請將以下內容截圖並寄送給開發人員以排除錯誤, 同時請確認您的檔案及內容是正確的\n'
      _msg += str(err) + "\n" + errinfo
      QMessageBox.critical(None, _G.MsgError, _msg, QMessageBox.Ok)

def dump_errorlog(err, errinfo):
  try:
    filename = _G.get_error_logname()
    if not filename:
      return False
    with open(filename, 'w') as fp:
      fp.write(f"{str(err)}\n{errinfo}\n{_G.get_witty_comment()}")
    return filename
  except Exception:
    return False

def flatten(ar):
  for i in ar:
    if isinstance(i, collections.Iterable) and not isinstance(i, str):
      for j in flatten(i):
        yield j
    else:
      yield i

def diff_string(a,b):
  return SequenceMatcher(None,a,b).ratio()