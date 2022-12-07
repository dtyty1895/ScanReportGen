from configparser import ConfigParser
from ast import literal_eval

FILE_ENCODING = 'utf8'

def load_all_config(file, merge=True):
  config = ConfigParser()
  config.optionxform = str 
  config.read(file, encoding=FILE_ENCODING)
  ret = {}
  if merge:
    for s in config.sections():
      for k in config[s].keys():
        ret[k] = convert_type(config[s].get(k))
  else:
    return config
  return ret

def convert_type(val):
  try:
    return literal_eval(val)
  except Exception:
    return literal_eval(f'"{val}"')

def change_config(file, config):
  writeback = load_all_config(file, False)
  for s in writeback.sections():
    for k in writeback[s].keys():
      if k in config:
        writeback[s][k] = config[k]
  try:
    with open(file, 'w', encoding=FILE_ENCODING) as fp:
      writeback.write(fp)
  except Exception:
    return False
  return True
