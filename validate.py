import streamlit as st
import numpy as np
import openpyxl
import pandas as pd
import calendar
import datetime, re
from dateutil.relativedelta import relativedelta
import requests
from requests.exceptions import RequestException, ConnectionError, HTTPError, Timeout
from io import BytesIO
from PIL import Image
import time
from urllib.parse import urlparse
import asyncio
import aiohttp
from aiohttp import ClientConnectorError, ClientResponseError, ServerTimeoutError

# 関数の定義
def kishuizon(x,col,exclusion):
  """
  機種依存をチェックする関数
  """
  try:
    word = str(x).encode('Shift_JIS')
    return ''
  except UnicodeEncodeError:
    error_words = []
    for i in range(len(str(x))):       
      if str(x)[i] in exclusion:
      # if str(x)[i] in ['～','－','①','②','③','④','⑤','⑥','⑦','⑧','⑨','⑩','♡']: # エラーに含めない文字列
        continue
      else:
        try:
          str(x)[i].encode('Shift_JIS')
        except UnicodeEncodeError:
          error_words.append(str(x)[i])
    if len(error_words) == 0:
      return ''
    else:
      return f'{col}列に機種依存文字がありますのでご注意ください：{error_words}'

def kaigyo(x,col):
  """
  セル内改行を検出する
  """
  # pattern = r'^[a-zA-Z0-9\-]+$'
  if '\n' in str(x):
    return f'{col}列にセル内改行'
  else:
    return ''

# def check_space(x,col):
#   """
#   連続した空白をチェックする関数
#   """
#   text = str(x)
#   if text.count(' ') > 2 or text.count('　') > 2:
#     return f'{col}列に空白が3つ以上含まれています'
#   else:
#     return ''
def check_space(x,col):
  """
  連続した空白をチェックする関数
  """
  text = str(x)
  if re.search(r' {3,}|\u3000{3,}', text):
    return f'{col}列に空白が3つ以上含まれています'
  else:
    return ''

def hankaku_eisu(x,col):
  """
  対象列が半角英数か判定する
  電話番号のハイフンは許容
  カンマも許容
  """
  # 空欄は無視
  if x == '' or pd.isna(x) or x == None or x == np.nan:
    return ''
  else:
    # if '緯度' in col or '経度' in col:
    #   pattern = r'^[a-zA-Z0-9\-\.]+$'
    # else:
    #   pattern = r'^[a-zA-Z0-9\-]+$'
    pattern = r'^[a-zA-Z0-9\-\.]+$'
    if bool(re.match(pattern, str(x))) == False:
      return f'{col}列に半角英数以外有'
    else:
      return ''

def mail_check(x,col):
  """
  対象列がメールアドレスとして有効か判定する
  """
  # 空欄は無視
  if x == '' or pd.isna(x) or x == None or x == np.nan:
    return ''
  pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
  if bool(re.match(pattern, str(x))) == False:
    return f'{col}列に無効な文字列'
  else:
    return ''


async def is_url_alive(url, col, session):
    """
    リンク切れかどうかを判定する関数
    リクエストを送信し、返ってきたコードによって出力を変化
    CORS制限にも対応
    """
    # 空欄は無視
    if url == '' or pd.isna(url) or url == None:
        return url, ''

    # ヘッダーの設定（CORS対応を追加）
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
        'Accept-Language': 'ja,en-US;q=0.7,en;q=0.3',
        'Origin': 'null',
        'Sec-Fetch-Site': 'cross-site',
        'Sec-Fetch-Mode': 'cors',
        'Sec-Fetch-Dest': 'empty',
        'Access-Control-Request-Method': 'GET',
        'Access-Control-Request-Headers': 'content-type',
    }

    try:
        # URLの正規化
        if not (url.startswith('http://') or url.startswith('https://')):
            url = 'https://' + url  # デフォルトでhttpsを試す

        # プリフライトリクエストのシミュレーション（OPTIONSリクエスト）
        try:
            async with session.options(url, timeout=5, headers=headers, allow_redirects=True, ssl=False) as options_response:
                if options_response.status == 403:
                    # OPTIONSリクエストが拒否された場合は、直接GETリクエストを試みる
                    pass
                elif options_response.status not in [200, 204]:
                    return url, f'{col}列でプリフライトリクエストエラー：{options_response.status}'
        except Exception:
            # プリフライトリクエストが失敗しても、GETリクエストは試行する
            pass

        # 実際のリクエストの実行（リダイレクトを許可）
        async with session.get(url, timeout=10, headers=headers, allow_redirects=True, ssl=False) as response:
            final_url = str(response.url)  # 最終的なURL
            status = response.status

            # CORSヘッダーのチェック
            cors_origin = response.headers.get('Access-Control-Allow-Origin')
            if cors_origin is not None and cors_origin != '*' and cors_origin != 'null':
                return url, f'{col}列でCORS制限により一部機能が制限される可能性があります'

            # ステータスコードの判定
            if status == 200:
                return url, ''
            elif status == 301 or status == 302 or status == 307 or status == 308:
                return url, f'{col}列でリダイレクトが発生しています（{status}）。新しいURL: {final_url}'
            else:
                return url, f'{col}列でHTTPエラー：{status}'

    except aiohttp.ClientSSLError:
        # SSL証明書エラーの場合、httpで再試行
        if url.startswith('https://'):
            try:
                http_url = 'http://' + url[8:]
                async with session.get(http_url, timeout=10, headers=headers, allow_redirects=True) as response:
                    if response.status == 200:
                        return url, ''
                    else:
                        return url, f'{col}列でSSL証明書エラー、HTTP接続も失敗：{response.status}'
            except Exception:
                return url, f'{col}列でSSL証明書エラー'
        return url, f'{col}列でSSL証明書エラー'

    except aiohttp.ClientConnectorError as e:
        if 'Cannot connect to host' in str(e):
            return url, f'{col}列でDNSエラー：ホストに接続できません'
        return url, f'{col}列でネットワーク接続エラー：{str(e)}'

    except aiohttp.ServerTimeoutError:
        return url, f"{col}列でタイムアウトエラー：サーバーの応答が遅いです"

    except aiohttp.ClientResponseError as e:
        if e.status == 401:
            return url, f'{col}列で認証エラー：{e.status}'
        elif e.status == 403:
            if 'CORS' in str(e) or 'Access-Control' in str(e):
                return url, f'{col}列でCORS制限によりアクセスが拒否されました'
            return url, f'{col}列でアクセス拒否：{e.status}'
        elif e.status == 404:
            return url, f'{col}列でページが存在しません：{e.status}'
        else:
            return url, f'{col}列でHTTPエラー：{e.status}'

    except Exception as e:
        if 'CORS' in str(e) or 'Access-Control' in str(e):
            return url, f"{col}列でCORS制限エラー：{str(e)}"
        return url, f"{col}列でその他のエラー：{str(e)}"




def extract_domain(url):
    """
    URLからドメインを抽出する。
    """
    # 引数がfloat型またはNoneであるかをチェック
    if isinstance(url, float) or url is None:
        return '', '-->無効なURL形式です（float型）'
    
    # URLが空文字列の場合
    if url == '':
        return '', ''
        
    try:
        # URLを文字列に変換して解析
        url = str(url)
        if not (url.startswith('http://') or url.startswith('https://')):
            url = 'http://' + url
            
        parsed_url = urlparse(url)
        
        # netloc（ネットワークロケーション部分）からドメインを取得
        domain = parsed_url.netloc
        if not domain:
            return '', '-->ドメインの抽出に失敗しました'
            
        # ドメインからサブドメインを取り除く
        domain_parts = domain.split('.')
        
        # JPドメインの特殊処理
        if domain_parts[-1] == 'jp':
            # ac.jp, ed.jp, go.jp などの場合は3つのパーツを保持
            if len(domain_parts) >= 3 and domain_parts[-2] in ['ac', 'ed', 'go', 'lg', 'ne', 'or', 'co']:
                if len(domain_parts) > 3:
                    domain = '.'.join(domain_parts[-3:])
                else:
                    domain = '.'.join(domain_parts)
            # その他のjpドメインは2つのパーツを保持
            else:
                if len(domain_parts) > 2:
                    domain = '.'.join(domain_parts[-2:])
                else:
                    domain = '.'.join(domain_parts)
        # 一般的なドメイン（.comなど）の処理
        else:
            if len(domain_parts) > 2:
                # www.example.com の場合はwwwを除去
                if domain_parts[0] == 'www':
                    domain = '.'.join(domain_parts[1:])
                else:
                    domain = '.'.join(domain_parts)
            else:
                domain = '.'.join(domain_parts)
                
        return domain, ''
    except Exception as e:
        return '', f'-->URLの解析に失敗しました: {str(e)}'

async def process_urls(df, web_col_name):
    # URLカラムを文字列型に変換
    df[web_col_name] = df[web_col_name].astype(str)
    
    web_result = {}
    # nanを除外してユニークなURLのリストを作成
    url_list = [url for url in list(set(df[web_col_name].unique())) if url != 'nan']
    
    domain_list = []
    now = datetime.datetime.now()
    log_txt = ''
    placeholder = st.empty()

    async with aiohttp.ClientSession() as session:
        log_txt += f'検証対象URL(対象{len(url_list)}件)\n'
        log_txt += f'{str(url_list)}\n'
        log_txt += '/' + '-'*20+'\n'
        log_txt += '# 検証開始'+'\n'
        for idx,url in enumerate(url_list):
            log_txt += f'プロセス　　：{idx+1}/{len(url_list)}\n'
            log_txt += f'重複ドメイン：{domain_list}\n'
            log_txt += f'検証URL　　 ：{url}\n'
            log_txt += f"検証開始時刻：{now.strftime('%H:%M:%S.%f')}\n"
            obj_domain, error_msg = extract_domain(url)
            log_txt += f'検証ドメイン：{obj_domain}\n'
            if error_msg:
                log_txt += f'{error_msg}\n'
                web_result[url] = 'URLの形式が無効です'
                log_txt += '/'+'-'*20+'\n'
                continue

            if obj_domain in domain_list:
                log_txt += f'-->ドメイン重複、2秒待機\n'
                await asyncio.sleep(2)
            
            if len(domain_list) >= 4:
                domain_list.pop(0)
            domain_list.append(obj_domain)

            result = await is_url_alive(url, web_col_name, session)
            log_txt += f"検証結果　　：{str(result[1])}\n"
            log_txt += '/'+'-'*20+'\n'
            if debug_mode:
                placeholder.code(log_txt,language='markdown')
            web_result[result[0]] = result[1]

    for url, e_word in web_result.items():
        df.loc[df[web_col_name]==url, 'エラー'] = df.loc[df[web_col_name]==url, 'エラー'] + ',' + e_word
    
    return df





def response_time(url):
  """
  リクエストの応答時間を出力
  """
  if url == '' or pd.isna(url) or url == None:
    return "対象外"
  else:
    try:
      response = requests.get(url)
      time_elapsed = response.elapsed.total_seconds()
      return time_elapsed
    except requests.exceptions.RequestException:
      return 'エラー'

def ido(x,col):
  """
  緯度の値が-90度から90度の範囲内にあるか判定する
  全角数字は半角に変換されて計算される
  """
  if pd.isna(x) or x == '': # 空欄の場合は何もしない
    return ''
  else:
    try:
      # if pd.isna(x):
      #   return '緯度が空欄です'
      # 浮動小数点数に変換し、範囲内かどうかをチェック
      latitude = float(x)
      if 20 <= latitude <= 46:
        return ''
      else:
        return '緯度の数値範囲外'
    except (ValueError, TypeError):
        return f'緯度に無効な文字列'
  
def keido(x,col):
  """
  経度の値が-180度から180度の範囲内にあるか判定する
  全角数字は半角に変換されて計算される
  """
  if pd.isna(x) or x == '': # 空欄の場合は何もしない
    return ''
  else:
    try:
      # if pd.isna(x):
      #   return '経度が空欄です'
      # 浮動小数点数に変換し、範囲内かどうかをチェック
      longitude = float(x)
      if 122 <= longitude <= 154:
        return ''
      else:
        return '経度の数値範囲外'
    except (ValueError, TypeError):
        return f'経度に無効な文字列'

def check_blank(x,col):
  """
  必須項目が入力されているかをチェックする関数
  """
  if pd.isna(x) or x == '':
    return f'{col}列が空欄です'
  else:
    return ''


def youbi(x, col):
  """
  カンマ区切りで曜日が入力されているか判定
  """
  if pd.isna(x) or x == '': # 空欄の場合は無視
    return ''
  else:
    valid_days = {'月', '火', '水', '木', '金', '土', '日', '祝'}
    try:
      days = set(x.split(','))  # カンマで分割してセットに変換
      if days.issubset(valid_days):
          return ''
      else:
          return f'曜日はカンマ区切りで入力してください'
    except AttributeError:
      return f'曜日はカンマ区切りで入力してください'

def is_time(value):
    """
    時間の形式が合っているかを確認する関数
    """
    return isinstance(value, (datetime.time,datetime.datetime))
def strTodatetime(time_val):
  """
  ## 施設用の関数（日付を含まない）
  時間列の値をdatetime.datetime形式へ変換する
  「xx:xx:xx」となっている形式は「xx:xx」へ揃える
  入力が「xx:xx」または「xx:xx:xx」以外の場合はエラーを返す
  """
  try:
    time_val = str(time_val)
    if len(time_val) > 5:
      time_val = datetime.datetime.strptime(time_val, '%H:%M:%S')
    elif len(time_val) <= 5:
      time_val = datetime.datetime.strptime(time_val, '%H:%M')
    return time_val
  except ValueError:
    return 'error'
def validate_business_hours(start, end):
  """
  "xx:xx" 形式でない場合、または開始時間が終了時間より後の場合にエラーを検出

  """
  if (pd.isna(start) or start == '') & (pd.isna(end) or end == ''):
    return ''
  else:
    if not is_time(strTodatetime(start)):
      errorword = '開始時間の形式が違います'
      if not is_time(strTodatetime(end)):
        errorword = '開始+終了時間の形式が違います'
        return errorword
      else:
        return errorword
    elif not is_time(strTodatetime(end)):
      return '終了時間の形式が違います'
    start = strTodatetime(start).strftime('%H:%M')
    end = strTodatetime(end).strftime('%H:%M')

    if (not (':' in start and ':' in end)) or (start == ':' or end == ':'):
        return '時間形式エラー'
    
    start_hours, start_minutes = map(int, start.split(':'))
    end_hours, end_minutes = map(int, end.split(':'))
    
    if start_hours >= end_hours or (start_hours == end_hours and start_minutes >= end_minutes):
        return '開始時間と終了時間が同じまたは逆転しています'
    
    return ''

def doreka_ireru(row, columns):
  limit = len(columns)
  if limit == 0:
    return ''
  num = 0
  for i in columns:
    if (row[i] is None) or (row[i] == '') or (row[i] == np.nan) or pd.isna(row[i]):
      num +=1
  if limit == num:
    return f'{columns}のうち、いずれかは入力'
  else:
    return ''

def maru(x,col):
  """
  指定した列に「○」または空欄以外の値が入っている場合を検出する関数
  """
  # 空欄は無視
  if x == '' or pd.isna(x) or x == None or x == np.nan:
    return ''
  if x is not None and x not in ['○','〇','','◯',1,0]:
    return f'{col}に無効な文字列'
  return ''

# 入力禁止列
def check_forbid(x,col):
  """
  入力禁止の列に値が入力されているかをチェックする関数
  """
  if pd.isna(x) or x == '':
    return ''
    
  else:
    return f'{col}列は入力禁止です'


# イベント専用関数
def event_time(start,end,event_day_range):
  """
  イベントの時間を検証する関数
  次の優先度で検出
  1. 時間形式が正しいか。yyyy-mm-ddの形式または、yyyy/mm/dd
  2. 開始時間と終了時間が逆転していないか
  3. 開始時間と終了時間が一緒の日であるか
  """
  # 両方空欄である場合は検出しない
  if (pd.isna(start) or start == '') & (pd.isna(end) or end == ''):
    return ''
  else:
    try:
        start,end = str(start), str(end)
        if len(start) > 19:
          start_time = datetime.datetime.strptime(start, '%Y-%m-%d %H:%M:%S.%f')
        else:
          start_time = datetime.datetime.strptime(start, '%Y-%m-%d %H:%M:%S')
        # start_time = datetime.datetime.strptime(start, '%Y-%m-%d %H:%M:%S.%f')
        if len(end) > 19:
          end_time = datetime.datetime.strptime(end, '%Y-%m-%d %H:%M:%S.%f')
        else:
          end_time = datetime.datetime.strptime(end, '%Y-%m-%d %H:%M:%S')
        # end_time = datetime.datetime.strptime(end, '%Y-%m-%d %H:%M:%S.%f')
        start_day, end_day = datetime.date(start_time.year,start_time.month,start_time.day),datetime.date(end_time.year,end_time.month,end_time.day)
        if start_day < event_day_range[0] or end_day > event_day_range[1]:
            return f'イベントが対象期間外です'
        elif start_time.day != end_time.day:
            return '開始日時と終了日時は同じ日にしてください'
        elif start_time >= end_time: 
            return '開始時間と終了時間が同じまたは逆転しています'
        else: # 時間が正しい
            return ''
    except ValueError:
        try:
            if len(start) > 19:
              start_time = datetime.datetime.strptime(start, '%Y/%m/%d %H:%M:%S.%f')
            else:
              start_time = datetime.datetime.strptime(start, '%Y/%m/%d %H:%M:%S')
            if len(end) > 19:
              end_time = datetime.datetime.strptime(end, '%Y/%m/%d %H:%M:%S.%f')
            else:
              end_time = datetime.datetime.strptime(end, '%Y/%m/%d %H:%M:%S')
            
            start_day, end_day = datetime.date(start_time.year,start_time.month,start_time.day),datetime.date(end_time.year,end_time.month,end_time.day)
            if start_day < event_day_range[0] or end_day > event_day_range[1]:
              return f'イベントが対象期間外です'
            elif start_time.day != end_time.day or start_time.month != end_time.month or start_time.year != end_time.year:
              return '開始日時と終了日時は同じ日にしてください'
            elif start_time >= end_time:
              return '開始時間と終了時間が同じまたは逆転しています'
            else: # 時間が正しい
                return ''
        except ValueError:
          return '開始時間または終了時間がフォーマットに従っていません'



def check_types(variable, check_list,err_text):
    """
    列の値が、指定したチェックリストの値に含まれない場合に検出
    """
    if pd.isna(variable) or str(variable) in list(check_list):
      return ''
    else:
      return err_text

def ex_error_check(file):
  """
  エクセルファイルを読み込むときに起動。
  エクセルのセルに数式が入っていた場合に検出。
  """
  ex_err_list = []
  wb = openpyxl.load_workbook(file)
  sheet = wb[wb.sheetnames[0]]
  for row in sheet.rows:
    for cell in row:
      if cell.data_type == 'f':
        ex_err_list.append(cell.coordinate)
  if len(ex_err_list) >0:
    st.error(f"ファイルの次のセルが数式のため、値に直してください{ex_err_list}")
    st.stop()

def check_word_in_list(lists,word):
  """
  リストに特定の文字が含まれている場合にTrueを返す
  """
  if len([s for s in lists if word in s])!=0:
    return True
  else:
    return False

# 場所と場所の名前
def basho_namae(row, columns):
  """
  場所と場所の名前両方に値が入っている場合に検出
  """
  limit = len(columns)
  num = 0
  for i in columns:
    if (row[i] is None) or (row[i] == '') or (row[i] == np.nan) or pd.isna(row[i]):
      num +=1
  if num == 0: # 両方値が入っている場合
    return f'{columns}は、いずれかのみ入力してください'
  else:
    return ''

# 重複する値のチェック
def find_duplicates(lst):
    """
    引数：リスト
    返り値：重複した値のリスト
    """
    seen = set()
    duplicates = set()
    for item in lst:
        if item in seen:
            duplicates.add(item)
        else:
            seen.add(item)
    return list(duplicates)
