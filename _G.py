from threading import Lock
from datetime import datetime
from collections import deque
import traceback
from random import random
import os
import configmanager

Version = "v2.2.0"
WindowTitle = '網頁掃描報告生成器'
MainWindow = None

FileTypeExcel = 'Excel Files (*.xlsx)'
FileTypeWord  = 'Word Files (*.doc *.docx)'
FileTypePPT   = 'PPT Files (*.ppt *.pptx)'
FileTypePDF   = 'PDF Files (*.pdf)'
FileTypeDir   = 'Folder (*./)'
Hostname      = ''

CacheFile = '.pycache'
READMEFile = "https://hackmd.io/irFVZ6yMRsCkD0-t7rrEBA?view"

# Generated file extenstion
GenFileExt = '.docx'

# Thread-messaging between worker and GUI thread
PipeMessages = []
PipeMutex    = Lock()
PipeError    = deque()
PipeWarning  = deque()
PipeSubInfo  = deque() # Info pipe to sub-thread

# 0: None, 1: +Error, 2: +Warning, 3: +Info, 4: +Debug
VerboseLevel = 4

# Object style in generated doc
DefaultDocTableStyle = 'List Table 3 Accent 1'
DefaultDocParagStyle = 'List Paragraph'
DocTableStyle = ''
DocParagStyle = ''
DocListTypeUnordered = "9"

# Keyword to find target information line(s)
KwordRiskCnt = 'Total alerts found'
KwordRiskRangeStart = ['Affected items'] 
KwordRiskRangeEnd   = ['Alerts details']
KwordAffectedItems = ['Affected items', 'Web Server', 'Details']
KwordMainhost = 'Start url'
KwordRisknameSuccessor = 'Classification'
KwordRiskDesc = 'Description'
KwordRiskImpact = 'Impact'
KwordRiskRecomd = 'Recommendation'
KwordRiskInfoEnd = ['References', 'Affected items']
KwordRiskLevel = 'Severity'
KwordDocDate = '日期'
KwordScanStartTime = 'Start time'
KwordExecutionDate = ''

KwordWebServer = 'Web Server'
RegexWebServer = r"^web\s*server$"
RegexAffectExpandStop = r"^request\s*header(s)?"
RegexURI = r"^http(s)?:\/\/"
RiskTagExpandAffectedItem = '###'

# Regex to match keyword (or its value)
RegexStartTime = r"(\d+)\-(\d+)-(\d+)"

# Excel keys
XlsDevSheetName = 'Developer'
XlsOwaspSheetName = 'OWASP'
XlsRiskColName  = '風險名稱'

# DOCuments appends to generated file
DocMeasureRangeTitles = ['監測範圍', '執行期間', '備註']
DocRiskCountTitles = [['風險等級', '高', '中', '低', '資訊風險'], ['數量']]
DocRiskListTitles = ['編號', '風險等級', '未通過項目', '數量']
DocRiskLevelTrans = {'High':'高', 'Medium':'中', 'Low':'低', 'Informational':'資訊風險'}
DocOwaspListTitle = ['編號', 'OWASP Top10', '未通過項目']
DocRiskDetails = ['風險性：', '風險內容概述：', '衝擊：', '影響範圍：', '建議：']

# Keyword REPlacement in template doc

RepRiskCntList = 'docRiskCntList'
RepRiskLevel = ["totalRiskCnt", "highRiskCnt", "midRiskCnt", "lowRiskCnt", "infoRiskCnt"]
RepRiskList = 'docRiskList'
RepOwaspList = 'docOwaspList'
RepOwaspRisk = 'docOwaspRisk'
RepRiskDesc = 'docRiskDescribe'
RepMeasureURL = 'measurementUrl'
RepCompanyName = 'OOOO'
RepCompanyAbbr = 'XXXX'
Repdate = 'YYYY/MM/DD'
# Font size in risk description section
RiskDescFontSize = 14

# Messages
MsgInit   = '資料讀取中...'
MsgHost   = '正在解析報告主機...'
MsgRisks  = '正在解析報告風險內容...'
MsgRiskRange = '正在解析報告風險範圍...'
MsgOwaList   = '正在解析 OWASP 列表...'
MsgOwaRisks  = '正在解析 OWASP 風險內容...'
MsgDocGen    = '正在生成報告文件...'
MsgUnwritable = '無法寫入到所選擇的存檔路徑, 請確認無非法字元且有權限存取!'
MsgError = '錯誤'
MsgInfo = '訊息'

MsgPipeWarnTargetOpened = '寫入的檔案已被其他程式開啟, 請關閉後再按下確認鍵以繼續儲存檔案'
MsgPipeContinue = '\x00\x50\x00CONTINUE'
MsgPipeStop  = "\x00\x50\x00STOP"
MsgPipeError = "\x00\x50\x00ERROR"
MsgPipeTerminated = "\x00\x50\x00TERMINATED"
MsgPipeRet = "\x00\x50\x00RET"
MsgPipeInfo = "\x00\x50\x00INFO"

UIInputFieldNames = ["Developer PDF", "OWASP PDF", "漏洞資料庫 Excel", "報告樣板 Word"]

WittyComments = [
  "幹　怎麼又掛了",
  "這不是 Bug, 這是 Feature :D",
  "以前只有我和上帝知道我的程式在寫啥，現在只有神知道了",
  "蛤?",
  "你以為是 Bug，但其實是我 DIO 噠！",
  "This is fine.",
]

# Config related
ConfigFile = 'config.cfg'
DefaultConfig = 'config.cfg.default'
Fonts = []
FontSizeMinMax  = [4, 128]

DOCX_PT_MAGIC = 12700

TargetVersion = 0

def em2pt(em):
  return em // DOCX_PT_MAGIC

def get_witty_comment():
  return "// " + WittyComments[int(random()*100) % len(WittyComments)]

def format_curtime():
  return datetime.strftime(datetime.now(), '%H:%M:%S')

def log_error(*args, **kwargs):
  if VerboseLevel >= 1:
    print(f"[{format_curtime()}] [ERROR]:", *args, **kwargs)

def log_warning(*args, **kwargs):
  if VerboseLevel >= 2:
    print(f"[{format_curtime()}] [WARNING]:", *args, **kwargs)

def log_info(*args, **kwargs):
  if VerboseLevel >= 3:
    print(f"[{format_curtime()}] [INFO]:", *args, **kwargs)

def log_debug(*args, **kwargs):
  if VerboseLevel >= 4:
    print(f"[{format_curtime()}] [DEBUG]:", *args, **kwargs)

def get_error_logname():
  folder = '錯誤日誌/'
  try:
    if not os.path.isdir('錯誤日誌'):
      os.mkdir('錯誤日誌')
  except Exception:
    folder = ''
  ret = f"{folder}errorlog_{str(datetime.now()).split('.')[0]}.log"
  _tr = str.maketrans({
    ':': '-', ' ': '_'
  })
  return ret.translate(_tr)

# Push message to pipe
def append_pipe_message(msg):
  global PipeMutex, PipeMessages
  PipeMutex.acquire()
  try:
    PipeMessages.append(msg)
  finally:
    PipeMutex.release()

# Pop messages from pipe
def pop_pipe_messages(cnt=0):
  global PipeMutex, PipeMessages
  PipeMutex.acquire()
  try:
    _msglen = len(PipeMessages)
    cnt = _msglen if cnt <= 0 or cnt > _msglen else cnt
    ret = PipeMessages[0:cnt]
    PipeMessages = PipeMessages[cnt:]
  finally:
    PipeMutex.release()
  return ret

## Load config
Config = {}

def load_config():
  global Config
  try:
    Config = configmanager.load_all_config(ConfigFile)
    log_info(f"Config loaded:\n{Config}")
    return True
  except Exception:
    log_error("Error while loading config, using default")
    Config = configmanager.load_all_config(DefaultConfig)
    return False

MsgInit   = '資料讀取中...'
MsgHostList   = '正在解析網站列表...'
MsgRisksCnt  = '正在解析報告風險數量...'
MsgHMRisk = '正在解析報告高中風險TOP10...'
MsgHighRiskSolve = '正在解析高風險弱點說明...'
MsgPPTGen    = '正在生成報告文件...'

contentname=[]

pdf_path=''

start_url=[]
high_risk =[]
mid_risk=[]
file_list=[]
risk = []

allFileList=0
range_value = 12
slide_add=0

HighWeak = {}
MidWeak = {}

ppt_params={
    'OOOO':"",
    'YYYY':"",
    'MM':"",
    'DD':"",
    'webCnt':''
}
ppt_table={
    'Test_target':'受測目標',
    'testRiskCnt':'測試結果風險數量統計',
    'highRiskWeak':"高風險弱點列表",
    'midRiskWeak':"中風險弱點列表",
    'high_risk_explain':'高風險弱點說明'
}

load_config()