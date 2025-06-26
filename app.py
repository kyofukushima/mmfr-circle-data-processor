import streamlit as st
import pandas as pd
import io
import time
import os
import chardet
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from copy import copy
from openpyxl.styles import Border, Side, Alignment, Font

# テンプレートファイルのパスを環境変数から取得（テスト時に切り替え可能）
TEMPLATE_FILE = os.getenv('TEMPLATE_FILE', 'template.xlsx')

def detect_encoding(file_content):
    """ファイルのエンコーディングを検出する"""
    result = chardet.detect(file_content)
    return result['encoding'], result['confidence']

def process_binary_columns(df):
    """0/1の値を持つ列の変換処理を行う
    
    Args:
        df (pd.DataFrame): 処理対象のデータフレーム
    
    Returns:
        tuple: (処理後のデータフレーム, 処理内容のデータフレーム)
    """
    # 変換対象の列を定義
    target_patterns = [
        '対象年齢',  # 対象年齢を含む列
        '要会費',
        '冊子掲載可',
        'HP掲載可',
        'オープンデータ掲載可'
    ]
    
    # 変換対象の列名を抽出
    target_columns = []
    for pattern in target_patterns:
        matched_columns = [col for col in df.columns if pattern in str(col)]
        target_columns.extend(matched_columns)
    
    # 重複を削除
    target_columns = list(dict.fromkeys(target_columns))
    
    # 処理内容をデータフレームとして作成
    process_df = pd.DataFrame(columns=['処理内容', '対象列'])
    if target_columns:
        for col in target_columns:
            # 列を文字列型に変換
            df[col] = df[col].astype(str)
            # 0を空欄に変換
            df[col] = df[col].replace('0', '')
            # 1を○に変換
            df[col] = df[col].replace('1', '○')
        
        # 処理内容のデータフレームを作成
        process_df = pd.DataFrame({
            '処理内容': ['0を空欄に変換および1を○に変換'],
            '対象列': [', '.join(target_columns)]
        })
    
    return df, process_df

def add_location_column(circle_data,df_f):
    """
    場所列の追加施設情報のエクスポートデータを抽出する
    S列「活動場所」の施設名称を参考に、施設情報のエクスポートデータからJ列「場所」を抽出・突合する
    J列「場所」から抽出・突合した情報をAY列の「場所」に入力する
    S列「活動場所」に施設名称がなかったり【育児サークル等地区表示用】○○区を指定している場合にはAY列の「場所」は空欄になるが、この場合空欄になるのが正。
    
    Returns:
        tuple: (処理後のデータフレーム, 処理内容のデータフレーム)
    """
    circle_data['場所'] = circle_data['活動場所'].map(df_f.set_index('施設名')['場所'])
    
    # 処理内容のデータフレームを作成
    process_df = pd.DataFrame({
        '処理内容': ['活動場所の施設名称から場所情報を抽出・突合'],
        '対象列': ['活動場所 → 場所']
    })
    
    return circle_data, process_df

def check_data_consistency(circle_data, last_month_data):
    """
    育児サークルデータと先月分データの整合性をチェックする
    
    Args:
        circle_data (pd.DataFrame): 育児サークルデータ
        last_month_data (pd.DataFrame): 先月分データ
    
    Returns:
        None
    
    Raises:
        st.stop(): データの不一致がある場合に処理を停止
    """
    # スラッグの重複チェック
    circle_duplicates = circle_data[circle_data['スラッグ'].duplicated()]['スラッグ'].unique()
    last_month_duplicates = last_month_data[last_month_data['スラッグ'].duplicated()]['スラッグ'].unique()
    
    error_messages = []
    
    if len(circle_duplicates) > 0:
        error_messages.append("### 育児サークルデータ内で重複しているスラッグ:")
        for slug in circle_duplicates:
            circle_names = circle_data[circle_data['スラッグ'] == slug]['サークル名'].tolist()
            error_messages.append(f"- スラッグ: {slug}")
            for name in circle_names:
                error_messages.append(f"  - サークル名: {name}")
    
    if len(last_month_duplicates) > 0:
        error_messages.append("\n### 先月分データ内で重複しているスラッグ:")
        for slug in last_month_duplicates:
            circle_names = last_month_data[last_month_data['スラッグ'] == slug]['サークル名'].tolist()
            error_messages.append(f"- スラッグ: {slug}")
            for name in circle_names:
                error_messages.append(f"  - サークル名: {name}")
    
    # スラッグの存在チェック
    circle_slugs = set(circle_data['スラッグ'])
    last_month_slugs = set(last_month_data['スラッグ'])
    
    # 育児サークルデータにのみ存在するスラッグ
    only_in_circle = circle_slugs - last_month_slugs
    # 先月分データにのみ存在するスラッグ
    only_in_last_month = last_month_slugs - circle_slugs
    
    if only_in_circle:
        error_messages.append("\n### 先月分データに存在しないスラッグ:")
        for slug in only_in_circle:
            circle_name = circle_data[circle_data['スラッグ'] == slug]['サークル名'].iloc[0]
            error_messages.append(f"- スラッグ: {slug} (サークル名: {circle_name})")
    
    if only_in_last_month:
        error_messages.append("\n### 育児サークルデータに存在しないスラッグ:")
        for slug in only_in_last_month:
            circle_name = last_month_data[last_month_data['スラッグ'] == slug]['サークル名'].iloc[0]
            error_messages.append(f"- スラッグ: {slug} (サークル名: {circle_name})")
    
    if error_messages:
        st.error("""
        ### データの不一致が検出されました
        
        {}
        
        ※ スラッグの重複や不一致を修正してから再度実行してください。
        """.format('\n'.join(error_messages)))
        st.stop()

def add_account_columns(circle_data, last_month_data):
    """
    先月分のデータからアカウント情報を追加する
    
    Args:
        circle_data (pd.DataFrame): 育児サークルデータ
        last_month_data (pd.DataFrame): 先月分のデータ
    
    Returns:
        tuple: (処理後のデータフレーム, 処理内容のデータフレーム)
    """
    # アカウント関連列の追加
    account_columns = ['ｱｶｳﾝﾄ発行有無', 'ｱｶｳﾝﾄ発行年月', 'アカウント発行の登録用メールアドレス']
    
    try:
        for col in account_columns:
            # スラッグをキーとしてマッピング
            mapping_dict = last_month_data.set_index('スラッグ')[col].to_dict()
            circle_data[col] = circle_data['スラッグ'].map(mapping_dict)
    except Exception as e:
        st.error(f"""
        アカウント情報の追加中にエラーが発生しました。
        エラー内容: {str(e)}
        
        以下を確認してください：
        1. スラッグに重複がないこと
        2. 必要な列（{', '.join(account_columns)}）が先月分データに存在すること
        """)
        st.stop()
    
    # 処理内容のデータフレームを作成
    process_df = pd.DataFrame({
        '処理内容': ['先月分データからアカウント情報を追加'],
        '対象列': [', '.join(account_columns)]
    })
    
    return circle_data, process_df

def validate_csv_file(csv_file):
    """CSVファイルの検証を行う"""
    # 基本的なエンコーディングリスト
    encodings = ['utf-8', 'shift-jis', 'cp932', 'euc-jp']
    detected_encoding = None
    debug_info = []
    
    # ファイルの内容を読み込む
    file_content = csv_file.read()
    csv_file.seek(0)
    
    # chardetによるエンコーディング検出
    detected_enc, confidence = detect_encoding(file_content)
    if detected_enc:
        encodings.insert(0, detected_enc)
        debug_info.append(f"chardetが検出したエンコーディング: {detected_enc} (信頼度: {confidence:.2f})")
    
    # 重複を削除
    encodings = list(dict.fromkeys(encodings))
    
    for encoding in encodings:
        try:
            debug_info.append(f"エンコーディング {encoding} で試行中...")
            
            # ファイルポインタを先頭に戻す
            csv_file.seek(0)
            
            # 最初の数行を読んでエンコーディングをチェック
            sample = file_content.decode(encoding)
            if not sample.strip():
                debug_info.append(f"  → ファイルが空です")
                continue
            
            # ファイルポインタを先頭に戻す
            csv_file.seek(0)
            
            # CSVとして読み込めるかチェック
            df = pd.read_csv(io.StringIO(sample), encoding=encoding)
            if df.empty:
                debug_info.append(f"  → データが空です")
                continue
            if len(df.columns) == 0:
                debug_info.append(f"  → 列が存在しません")
                continue
                
            # ファイルポインタを先頭に戻す
            csv_file.seek(0)
            detected_encoding = encoding
            debug_info.append(f"  → 正常に読み込めました")
            return df, detected_encoding, debug_info
            
        except UnicodeDecodeError as e:
            debug_info.append(f"  → デコードエラー: {str(e)}")
            continue
        except pd.errors.EmptyDataError:
            debug_info.append(f"  → 空のCSVファイル")
            raise ValueError("CSVファイルが空です")
        except Exception as e:
            debug_info.append(f"  → その他のエラー: {str(e)}")
            continue
    
    error_msg = "CSVファイルのエンコーディングを認識できません。以下のいずれかの形式で保存してください：UTF-8、Shift-JIS、CP932、EUC-JP"
    if st.session_state.get('debug_mode', False):
        error_msg += "\n\nデバッグ情報:\n" + "\n".join(debug_info)
    raise ValueError(error_msg)

def copy_cell_format(source_cell, target_cell):
    """セルの書式をコピーする"""
    if source_cell.has_style:
        target_cell.font = copy(source_cell.font)
        target_cell.border = copy(source_cell.border)
        target_cell.fill = copy(source_cell.fill)
        target_cell.number_format = copy(source_cell.number_format)
        target_cell.protection = copy(source_cell.protection)
        target_cell.alignment = copy(source_cell.alignment)

def validate_excel_file(excel_file):
    """先月分のExcelファイルの検証と読み込みを行う
    
    Args:
        excel_file: アップロードされたExcelファイル
    
    Returns:
        pd.DataFrame: 読み込んだデータフレーム
    """
    try:
        # Excelファイルを読み込む（2,3行目をスキップ）
        df = pd.read_excel(excel_file, skiprows=[1,2])
        
        # 基本的な検証
        if df.empty:
            raise ValueError("Excelファイルにデータが存在しません")
            
        if len(df.columns) == 0:
            raise ValueError("Excelファイルに列が存在しません")
        
        # ヘッダーの存在確認
        if df.columns.isna().any():
            raise ValueError("ヘッダー行に空の列名が存在します")
        
        return df
        
    except pd.errors.EmptyDataError:
        raise ValueError("Excelファイルが空です")
    except Exception as e:
        raise ValueError(f"Excelファイルの読み込み中にエラーが発生しました: {str(e)}")

def hide_columns(worksheet):
    """特定の列を非表示にする
    
    Args:
        worksheet: 対象のワークシート
    """
    # 非表示にする列名のリスト
    columns_to_hide = [
        'スラッグ',
        'ステータス',
        '参加者の条件(妊娠後半)',
        '参加者の条件(出産)',
        '参加者の条件(1歳後半)',
        '参加者の条件(2歳後半)',
        '申込方法備考',
        '活動日_営業時間ラベル',
        '活動日_営業曜日ラベル',
        '代表者',
        '団体名'
    ]
    
    # ヘッダー行から列のインデックスを取得
    header_row = 1  # ヘッダーは1行目にある
    for column in worksheet.iter_cols(min_row=header_row, max_row=header_row):
        if column[0].value in columns_to_hide:
            col_letter = get_column_letter(column[0].column)
            worksheet.column_dimensions[col_letter].hidden = True

def add_borders(worksheet, start_row, end_row, start_col, end_col):
    """データ範囲に枠線を追加する
    
    Args:
        worksheet: 対象のワークシート
        start_row: 開始行（1始まり）
        end_row: 終了行（1始まり）
        start_col: 開始列（1始まり）
        end_col: 終了列（1始まり）
    """
    # 枠線のスタイルを定義
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # 指定範囲の各セルに枠線を設定
    for row in range(start_row, end_row + 1):
        for col in range(start_col, end_col + 1):
            cell = worksheet.cell(row=row, column=col)
            cell.border = thin_border

def find_data_range(worksheet):
    """データが存在する範囲を特定する
    
    Args:
        worksheet: 対象のワークシート
    
    Returns:
        tuple: (最終行, 最終列)
    """
    max_row = 1
    max_col = 1
    
    # 最終行を特定
    for row in worksheet.iter_rows():
        if any(cell.value is not None for cell in row):
            max_row = row[0].row
    
    # 最終列を特定
    for col in worksheet.iter_cols():
        if any(cell.value is not None for cell in col):
            max_col = col[0].column
    
    return max_row, max_col

def set_row_height_and_format(worksheet, start_row, end_row, height=20):
    """行の高さを設定し、セルの書式を設定する
    
    Args:
        worksheet: 対象のワークシート
        start_row: 開始行（1始まり）
        end_row: 終了行（1始まり）
        height: 行の高さ（デフォルト: 20）
    """
    # セルの書式設定（折り返し有効、左揃え）
    alignment = Alignment(
        wrap_text=True,  # 折り返し
        horizontal='left',  # 左揃え
        vertical='center'  # 縦方向は中央揃え
    )
    
    # フォント設定
    font = Font(
        name='メイリオ',  # フォント名
        size=12,         # フォントサイズ
    )
    
    # 指定範囲の各行に対して設定
    for row in range(start_row, end_row + 1):
        # 行の高さを設定
        worksheet.row_dimensions[row].height = height
        
        # その行の各セルの書式を設定
        for cell in worksheet[row]:
            cell.alignment = alignment
            cell.font = font

def setup_conditional_formatting(worksheet):
    """条件付き書式を設定する
    
    Args:
        worksheet: 対象のワークシート
    """
    from openpyxl.formatting.rule import Rule
    from openpyxl.styles import PatternFill
    from openpyxl.styles.differential import DifferentialStyle
    
    # 色の定義
    red_fill = PatternFill(start_color='FFFF0000', end_color='FFFF0000', fill_type='solid')
    yellow_fill = PatternFill(start_color='FFFFFF00', end_color='FFFFFF00', fill_type='solid')
    green_fill = PatternFill(start_color='FF00FF00', end_color='FF00FF00', fill_type='solid')
    
    # スタイルの定義
    red_style = DifferentialStyle(fill=red_fill)
    yellow_style = DifferentialStyle(fill=yellow_fill)
    green_style = DifferentialStyle(fill=green_fill)
    
    # 条件付き書式のリスト
    conditional_rules = [
        {
            'name': 'スラッグの差分検出',
            'description': 'スラッグが空、または入力されているがoriginalに見つからないものを検出',
            'formula': 'OR($B1="",ISERROR(MATCH($B1,INDIRECT("original!B1:B1048576"),0)))',
            'range': 'B1:B1048576',
            'style': red_style
        },
        # 追加の条件付き書式はここに追加
                 {
             'name': '変更箇所の検出',
             'description': '同じスラッグを持つ行の同じ列のセルを比較　⇒ 該当するセルだけ黄色く着色',
             'formula': 'A1<>INDIRECT("original!"&ADDRESS(MATCH($B1,INDIRECT("original!B1:B1048576"),0),COLUMN(),4,1))',
             'range': 'A1:ZZ1048576',
             'style': yellow_style
         },
         {
             'name': '追加行の検出',
             'description': '入力されているスラッグがoriginalに見つからないおよびサークル名がoriginalに見つからない',
             'formula': 'OR(ISERROR(MATCH($B1,INDIRECT("original!B1:B1048576"),0)),ISERROR(MATCH($C1,INDIRECT("original!C1:C1048576"),0)))',
             'range': 'A1:ZZ1048576',
             'style': green_style
         }
    ]
    
    # 条件付き書式を適用
    for rule_config in conditional_rules:
        rule = Rule(
            type="expression",
            formula=[rule_config['formula']],
            stopIfTrue=True,
            dxf=rule_config['style']
        )
        worksheet.conditional_formatting.add(rule_config['range'], rule)
        
        # デバッグモード時に設定内容を出力
        if st.session_state.get('debug_mode', False):
            st.info(f"条件付き書式を設定: {rule_config['name']} - {rule_config['description']}")

def process_files(circle_data, facility_data=None, last_month_data=None):
    """Pandasを使用したファイル処理"""
    start_time = time.time() if st.session_state.get('debug_mode', False) else None
    
    # 処理内容を記録するデータフレームを作成
    process_df = pd.DataFrame(columns=['処理内容', '対象列'])
    
    # 0/1の値を持つ列の変換処理
    circle_data, binary_process_df = process_binary_columns(circle_data)
    if not binary_process_df.empty:
        process_df = pd.concat([process_df, binary_process_df], ignore_index=True)
    
    # 場所列の追加
    circle_data, location_process_df = add_location_column(circle_data,facility_data)
    process_df = pd.concat([process_df, location_process_df], ignore_index=True)
    
    # アカウント情報の追加
    circle_data, account_process_df = add_account_columns(circle_data, last_month_data)
    process_df = pd.concat([process_df, account_process_df], ignore_index=True)
    
    # 処理内容の表示
    if not process_df.empty:
        with st.expander("処理内容を確認する"):
            st.dataframe(process_df, use_container_width=True, hide_index=True)
    
    # 処理後のデータフレームを表示
    with st.expander("処理後のデータフレームを確認する"):
        st.dataframe(circle_data, use_container_width=True)
    
    # ファイルを保存
    output = io.BytesIO()
    template_wb = load_workbook(TEMPLATE_FILE)
    template_ws = template_wb.active
    
    # シート名を'original'に変更
    template_ws.title = 'original'
    
    # テンプレートの内容をそのままコピー
    template_wb.save(output)
    output.seek(0)
    
    # 保存したファイルを再度開く
    wb = load_workbook(output)
    original_ws = wb['original']
    
    # CSVデータを書き込む（ヘッダー行を除く）
    if len(circle_data) > 0:  # データが存在する場合のみ処理
        # CSVの列数がテンプレートの列数を超えていないかチェック
        if len(circle_data.columns) > template_ws.max_column:
            raise ValueError(f"CSVファイルの列数（{len(circle_data.columns)}列）がテンプレートの列数（{template_ws.max_column}列）を超えています。")
        
        # データを一括で書き込む
        data_values = circle_data.values
        for row_idx, row in enumerate(data_values, start=4):  # 4行目から開始
            for col_idx, value in enumerate(row, start=1):
                cell = original_ws.cell(row=row_idx, column=col_idx)
                cell.value = value
                # テンプレートの同じ位置のセルから書式をコピー
                template_cell = template_ws.cell(row=row_idx, column=col_idx)
                copy_cell_format(template_cell, cell)
        
        # データが存在する範囲を特定
        max_row, max_col = find_data_range(original_ws)
        
        # データ部分に枠線を追加（1行目から最終行まで）
        add_borders(original_ws, 1, max_row, 1, max_col)
        
        # 行の高さとセル書式を設定（4行目からデータ最終行まで）
        set_row_height_and_format(original_ws, 4, max_row)
    
    # 特定の列を非表示にする
    hide_columns(original_ws)
    
    # シートをコピーして'circle_info'シートを作成
    circle_info_ws = wb.copy_worksheet(original_ws)
    circle_info_ws.title = 'circle_info'
    
    # originalシートを非表示にする
    original_ws.sheet_state = 'hidden'
    
    # 条件付き書式の設定
    setup_conditional_formatting(circle_info_ws)
    
    # シートのグループを解除
    for ws in wb.worksheets:
        ws.sheet_view.tabSelected = False
    
    # アクティブシートを明示的に設定（circle_infoシートをアクティブに）
    wb.active = circle_info_ws
    
    # ファイルを保存
    output.seek(0)
    wb.save(output)
    
    processing_time = time.time() - start_time if st.session_state.get('debug_mode', False) else None
    
    output.seek(0)
    return output, processing_time

def initialize_session_state():
    """セッション状態の初期化"""
    if 'debug_mode' not in st.session_state:
        st.session_state.debug_mode = False

def validate_order_column(df):
    """順番列の値を検証する
    
    Args:
        df (pd.DataFrame): 検証対象のデータフレーム
    
    Raises:
        ValueError: 順番列に数値以外の値が含まれている場合
    """
    if '順番' not in df.columns:
        return
    
    # 数値以外の値を含む行を検出
    non_numeric_rows = df[pd.to_numeric(df['順番'], errors='coerce').isna()]
    if not non_numeric_rows.empty:
        error_message = ["### エラー: 「順番」列に数値以外の値が含まれています"]
        for _, row in non_numeric_rows.iterrows():
            error_message.append(f"- サークル名: {row['サークル名']}")
            error_message.append(f"  - スラッグ: {row['スラッグ']}")
            error_message.append(f"  - 順番: {row['順番']}")
        
        raise ValueError("\n".join(error_message))
    
    # 1未満の値を含む行を検出
    invalid_rows = df[pd.to_numeric(df['順番']) < 1]
    if not invalid_rows.empty:
        warning_message = ["### 警告: 「順番」列に1未満の値が含まれている行があります"]
        for _, row in invalid_rows.iterrows():
            warning_message.append(f"- サークル名: {row['サークル名']}")
            warning_message.append(f"  - スラッグ: {row['スラッグ']}")
            warning_message.append(f"  - 順番: {row['順番']}")
        
        st.warning("\n".join(warning_message))

def main():
    initialize_session_state()
    
    # サイドバーにデバッグモードの切り替えを追加
    with st.sidebar:
        st.session_state.debug_mode = st.checkbox("デバッグモード", value=st.session_state.debug_mode)
        
        # バージョン情報（控えめに表示）
        st.markdown("")
        st.caption("v1.0 - 2025/06/26")
    
    st.title("育児サークル情報処理アプリ")
    
    # デバッグモード時のみ表示される情報
    if st.session_state.debug_mode:
        st.write("### デバッグ情報")
        st.write("デバッグモードが有効です")
    
    st.write("### ファイルのアップロード")
    
    # 育児サークルCSVファイルのアップロード
    csv_file = st.file_uploader("育児サークルCSVファイルを選択してください", type=['csv'])
    if csv_file:
        try:
            # CSVファイルの検証と読み込み
            circle_data, encoding, debug_info = validate_csv_file(csv_file)
            
            # 順番列の検証（検証の必要性について確認中。必要であればコメントアウト解除）
            # validate_order_column(circle_data)
            
            st.success("育児サークルCSVファイルが正常に読み込まれました")
            with st.expander("育児サークルデータを確認する"):
                st.dataframe(circle_data, use_container_width=True)
        except ValueError as e:
            st.error(f"育児サークルCSVファイルのエラー: {str(e)}")
        except Exception as e:
            st.error(f"育児サークルCSVファイルの予期せぬエラー: {str(e)}")
    
    # 施設情報CSVファイルのアップロード
    facility_csv_file = st.file_uploader("施設情報CSVファイルを選択してください", type=['csv'])
    if facility_csv_file:
        try:
            # 施設情報CSVファイルの検証と読み込み
            facility_data, facility_encoding, facility_debug_info = validate_csv_file(facility_csv_file)
            st.success("施設情報CSVファイルが正常に読み込まれました")
            with st.expander("施設情報データを確認する"):
                st.dataframe(facility_data, use_container_width=True)
        except ValueError as e:
            st.error(f"施設情報CSVファイルのエラー: {str(e)}")
        except Exception as e:
            st.error(f"施設情報CSVファイルの予期せぬエラー: {str(e)}")
    
    # 先月分のデータ（Excelファイル）のアップロード
    last_month_file = st.file_uploader("先月分のデータ（Excelファイル）を選択してください", type=['xlsx'])
    if last_month_file:
        try:
            # Excelファイルの検証と読み込み
            last_month_data = validate_excel_file(last_month_file)
            
            # データの整合性チェック（スラッグの一致確認）
            if 'circle_data' in locals() and circle_data is not None:
                check_data_consistency(circle_data, last_month_data)
            
            st.success("先月分のExcelファイルが正常に読み込まれました")
            with st.expander("先月情報データを確認する"):
                st.dataframe(last_month_data, use_container_width=True)
        except ValueError as e:
            st.error(f"先月分のExcelファイルのエラー: {str(e)}")
        except Exception as e:
            st.error(f"先月分のExcelファイルの予期せぬエラー: {str(e)}")
    
    # 全てのデータが揃っているか確認
    all_data_ready = (
        'circle_data' in locals() and circle_data is not None and
        'facility_data' in locals() and facility_data is not None and
        'last_month_data' in locals() and last_month_data is not None
    )
    
    if all_data_ready:
        st.success("全てのファイルが正常に読み込まれました。処理を開始できます。")
        
        # 自治体名の入力フィールドを追加（デフォルト値：北九州市様）
        municipality = st.text_input("自治体名", value="北九州市様", help="ダウンロードファイル名に使用される自治体名を入力してください")
        
        if st.button("処理開始"):
            try:
                # ファイル処理を実行
                output, proc_time = process_files(
                    circle_data,
                    facility_data=facility_data,
                    last_month_data=last_month_data
                )
                
                # デバッグモード時のみ処理時間と行数を表示
                if st.session_state.get('debug_mode', False):
                    st.info(f"処理時間: {proc_time:.3f}秒")
                    template_wb = load_workbook(TEMPLATE_FILE)
                    template_ws = template_wb.active
                    st.info(f"処理したデータ行数: {len(circle_data)-1}行")  # ヘッダーを除く
                    st.info(f"CSVファイルの列数: {len(circle_data.columns)}列")
                    st.info(f"テンプレートファイルの列数: {template_ws.max_column}列")
                
                # 現在の月を取得
                from datetime import datetime
                current_month = datetime.now().month
                
                # ファイル名を生成
                file_name = f"【{municipality}】育児サークル等修正用データ（{current_month}月分）.xlsx"
                
                # ダウンロードボタンを表示
                st.download_button(
                    label="処理済みファイルをダウンロード",
                    data=output,
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                st.success("ファイルの処理が完了しました！")
                
            except ValueError as e:
                st.error(str(e))
            except Exception as e:
                st.error(f"予期せぬエラーが発生しました: {str(e)}")
    
    st.write("### 使い方")
    st.write("1. CSVファイルをアップロードしてください")
    st.write("2. 「処理開始」ボタンをクリックしてください")
    st.write("3. 処理が完了したら、「処理済みファイルをダウンロード」ボタンが表示されます")

if __name__ == "__main__":
    main() 