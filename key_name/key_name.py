"""[]

  """

# prtfos venv

##

import openpyxl as pyxl
import pandas as pd
from pathlib import PurePath

import re
from persiantools import characters , digits
from openpyxl.utils.dataframe import dataframe_to_rows
from copy import copy

fp = '/Users/mahdimir/Dropbox/keyData/key_name/key_name.xlsx'

class Cols :
  def __init__(self) :
    self.key = 'key'
    self.cn = 'companyName'
    self.isCnUniq = 'isCompanyNameUnique'
    self.isKeyEqPrevKey = 'isKeyEqPrevKey'

cols = Cols()

def apply_func_on_notna_rows_of_cols(idf , cols , func) :
  for col in cols :
    _ms = idf[col].notna()
    idf.loc[_ms , col] = idf.loc[_ms , col].apply(func)
  return idf

def normalize_str(ist: str) :
  os = characters.ar_to_fa(str(ist))
  os = digits.ar_to_fa(os)
  os = digits.fa_to_en(os)
  os = os.strip()
  replce = {
      (r"\u202b" , None)          : ' ' ,
      (r'\u200c' , None)          : ' ' ,
      (r'\u200d' , None)          : '' ,
      (r'ء' , None)               : ' ' ,
      (r':' , None)               : ' ' ,
      (r'\s+' , None)             : ' ' ,
      (r'آ' , None)               : 'ا' ,
      (r'أ' , None)               : 'ا' ,
      (r'ئ' , None)               : 'ی' ,
      (r'\bETF\b' , None)         : '' ,
      (r'\(سهامی\s*عام\)' , None) : '' ,
      (r'\bسهامی\s*عام\b' , None) : '' ,
      (r'[\(\)]' , None)          : '' ,
      (r'\.+$' , None)            : '' ,
      (r'^\.+' , r'^\.+\d+$')     : '' ,
      }
  for key , val in replce.items() :
    if key[1] is not None :
      if not re.match(key[1] , os) :
        os = re.sub(key[0] , val , os)
    else :
      os = re.sub(key[0] , val , os)
  os = os.strip()
  return os

def define_wos_cols(df , which_cols: list) :
  for col in which_cols :
    msk = df[col].notna()
    ncol = col + '_wos'
    _pre = df.loc[msk , col]
    df.loc[msk , ncol] = _pre.str.replace(r'\s' , '' , regex=True)
  return df

def excel_style(row , col) :
  """ Convert given row and column number to an Excel-style cell name. """
  LETTERS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
  result = []
  while col :
    col , rem = divmod(col - 1 , 26)
    result[:0] = LETTERS[rem]
  return ''.join(result) + str(row)

def make_styles_alike(ws0 , ws1) :
  ws1.conditional_formatting = ws0.conditional_formatting

  for col in ws0.columns :
    wid = ws0.column_dimensions[col[0].column_letter].width
    ws1.column_dimensions[col[0].column_letter].width = wid

  title_style = ws0['A1'].style
  title_font = copy(ws0['A1'].font)

  for col in ws1.columns :
    col[0].style = title_style
    col[0].font = title_font

  for col in ws1.columns :
    col_1st_cell_add = col[0].column_letter + '2'
    _add = col_1st_cell_add
    _style = ws0[_add].style
    _font = copy(ws0[_add].font)

    for cell in col[1 :] :
      cell.style = _style
      cell.font = _font

  ws1.freeze_panes = 'A2'

  return ws1

def main() :
  pass

  ##
  wb0 = pyxl.load_workbook(fp)

  ##
  ws0 = wb0.active

  ##
  init_cols = [x[0].value for x in ws0.columns]

  ##
  df = pd.DataFrame(ws0.values , columns=init_cols)
  df = df.iloc[1 : , :]

  ##
  wb0.close()

  ##
  cols_list = [cols.key , cols.cn]
  df = apply_func_on_notna_rows_of_cols(df , cols_list , normalize_str)

  ##
  df = define_wos_cols(df , cols_list)

  ##
  df = df.drop_duplicates(subset=[x + '_wos' for x in cols_list])

  ##
  cols_list_1 = [cols.key, cols.cn]

  msk = df[cols_list_1].isna().any(axis=1)
  msk |= df[cols_list_1].eq('').any(axis=1)

  df1 = df[msk]
  df = df[~ msk]

  ##
  fnds_1 = df[init_cols]

  ##
  df = fnds_1

  ##
  df = df.sort_values(by=cols.key)

  ##
  wb1 = pyxl.Workbook()
  ws1 = wb1.active

  ##
  df = df.fillna(value='')

  for r in dataframe_to_rows(df , index=False , header=True) :
    ws1.append(r)

  ##
  # =IF(OR(COUNTIF(A: A , A2)=1 , A2 = "") , TRUE , FALSE)
  compn_letter_i = init_cols.index(cols.cn)
  compn_letter = excel_style(1 , compn_letter_i + 1)[:1]

  iscompn_uniq_letter_i = init_cols.index(cols.isCnUniq)
  iscompn_uniq_letter = excel_style(1 , iscompn_uniq_letter_i + 1)[:1]

  ##
  coc = compn_letter
  for i in range(2 , len(df) + 2) :
    celladd = coc + str(i)
    val = f'=IF(OR(COUNTIF({coc}:{coc},{celladd})=1,{celladd}=""),TRUE,FALSE)'
    ws1[iscompn_uniq_letter + str(i)] = val

  ##
  key_i = init_cols.index(cols.key)
  key_l = excel_style(1 , key_i + 1)[:1]

  is_key_eq_i = init_cols.index(cols.isKeyEqPrevKey)
  is_key_eq_l = excel_style(1 , is_key_eq_i + 1)[:1]

  ##
  for i in range(2 , len(df) + 2) :
    val = f'={key_l}{i}={key_l}{i-1}'
    ws1[is_key_eq_l + str(i)] = val

  ##
  ws1 = make_styles_alike(ws0 , ws1)

  ##
  wb1.save(fp)

  ##
  wb1.close()

##


if __name__ == "__main__" :
  main()
  print(f'{PurePath(__file__).name} Done.')

##

