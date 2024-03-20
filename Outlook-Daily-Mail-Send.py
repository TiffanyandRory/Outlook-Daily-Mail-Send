import pandas as pd
from datetime import datetime
import win32com.client
import xlwings as xw
import time
import re
from PIL import ImageGrab
import os
import gc

# 清理數據
def clean_newlines(df):
    return df.replace(r'\n', ' ', regex=True)

# 將Dashboard小數點轉%
def format_cell_value(x, r, c):
    try:
        if pd.isna(x): # chk值是否為NaN
            return " "  # 如果是NaN，返回空字符串
        float_val = float(x) # 特例 -> 如果數值為1，則r/n 100%
        if float_val == 1 and 5 <= r <= 9 and c == 3: 
             #Machine utilization with PO% 行 而已
             return "100%"
        if float_val.is_integer(): # chk值是否為整数
            return str(int(float_val)).replace('\n', ' ')  # 如果是整數，r/n原值
        else:
            return f'{float_val:.2%}'.replace('\n', ' ') # 是小數，轉%+小術後兩位
    except ValueError: # if r/n失敗，返回原值
        return x.replace('\n', ' ')

def convert_data_to_html_table(df):
    # 處理特定區域
    def format_specific_cell(r, c, value):
        if r == 5 and 2 <= c <= 8:  # iloc[5, 2:9] 對應於第6行，第3到第10列的單元格
            return f'<td style="font-weight: bold; color: blue; font-size: 20px;">{format_cell_value(value, r, c)}</td>'
        else:
            return f'<td style="font-size: 20px;">{format_cell_value(value, r, c)}</td>'
    
    # 建HTML表格
    sections = {}
    table_html = '<table style="background-color: #ffd34c; width: 200%; text-align: center;" class="centered">'
    for i, row in df.iloc[4:10, 2:9].iterrows():
        table_html += '<tr>'
        for j, value in enumerate(row):
            table_html += format_specific_cell(i, j+2, value)
        table_html += '</tr>'
    table_html += '</table>'

    sections['d6_l7'] = table_html

    html_parts = ['<p>{}</p>'.format(sections[key]) for key in sections]
    return ''.join(html_parts)


def generate_html_from_selected_data(df):
    column_styles = {
        'Group': 'color: blue; background-color: #ffd34c;',
        '區域代碼': 'color: blue; background-color: #ffd34c;',
        '機台號碼': 'color: blue; background-color: #ffd34c;',
        '目前設備狀況': 'color: blue; background-color: #ffd34c;',
        '生產狀況類別': 'color: blue; background-color: #ffd34c;',
        '專案名稱': 'color: blue; background-color: #ffd34c;',
        '模號': 'color: blue; background-color: #ffd34c;',
        '品名': 'color: blue; background-color: #ffd34c;',
        'priority': 'color: blue; background-color: #ffd34c;',
        '生產狀況敘述': 'color: blue; background-color: #ffd34c;',
        '備註': 'color: blue; background-color: #ffd34c;',
        'dcc週秒': 'color: blue; background-color: #ffd34c;',
        '實際週秒': 'color: blue; background-color: #ffd34c;',
        '週秒差異備註': 'color: blue; background-color: #ffd34c;',
        'sizemachine': 'color: blue; background-color: #ffd34c;',
        'screw': 'color: blue; background-color: #ffd34c;',
        'brands': 'color: blue; background-color: #ffd34c;',
        'tooling_tool_no': 'color: blue; background-color: #ffd34c;'
    }

    html = '<table border="1" style="border-collapse: collapse; width: 100%;" class="centered">'
    html += '<tr>' #+表頭
    for col in df.columns:
        style = column_styles.get(col, '') # 取目前列的樣式，如果没有樣式r/n空字符
        if style: # 依據樣式是否na + 不同樣式表頭
            html += f'<th style="{style}">{col}</th>'
        else:
            html += f'<th>{col}</th>'
    html += '</tr>'

    for _, row in df.iterrows():
        html += '<tr>'
        for value in row:
            # td 首件中 開機中 模具異常或修模 CSS
            if value == '首件中':
                value = f'<td style="color: darkorange; font-size: 18px;">{value}</td>'
            elif value == '開機中':
                value = f'<td style="color: dodgerblue; font-size: 18px;">{value}</td>'
            elif value == '模具異常或修模':
                value = f'<td style="color: firebrick; font-size: 18px;">{value}</td>'
            else:
                value = f'<td style="font-size: 18px;">{value}</td>'
            html += value
        html += '</tr>'
    html += '</table>'
    return html


def send_email(subject, body, image_path, body2, body3, recipients):
    try:
        outlook = win32com.client.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)
        mail.Subject = subject

        html_body = f"""
        <html>
        <head>
    
        </head>
        <body>
            <h2>Dashboard:</h2>
            {body}
            <h2>Chart:</h2>
            <img src="{image_path}" alt="請使用電腦設備觀看(已連上內網的設備)，並確保您具有 權限">
            <h2>停機總數量:</h2>
            {body2}
            <h2>大噸數機台開機狀況:</h2>
            {body3}
        </body>
        </html>
        """
        mail.HTMLBody = html_body
        mail.To = recipients
        mail.Send()
        print("Email sent successfully!")
    except Exception as e:
        print("Failed to send email:", str(e))

def save_excel_range_as_image(excel_path, sheet_name, range, output_directory):
    excel = win32com.client.Dispatch("Excel.Application")
    workbook = excel.Workbooks.Open(excel_path)
    sheet = workbook.Sheets(sheet_name)  # 指定工作表
    sheet.Range(range).CopyPicture(Format=win32com.client.constants.xlBitmap)
    
    image = ImageGrab.grabclipboard()

    if not os.path.exists(output_directory):
        os.makedirs(output_directory)

    current_date = datetime.now().strftime("%Y-%m-%d")

    # 輸出路徑&文件名
    output_path = os.path.join(output_directory, f"daily_mail_image_{current_date}.png")

    image.save(output_path)
    workbook.Close(False)
    excel.Quit()
    return output_path

def main():
    file_path = r'file location'
    output_directory = r'chart location'

    # Excel VBA運行
    app = xw.App(visible=True)
    book = app.books.open(file_path)
    app.macro('refresh_dashboard')()
    time.sleep(60)  #30S 應該就夠，但60S比較保險
    book.save()
    book.close()
    app.quit()
    
    gc.collect() # 釋放內存記憶體

    #其他VBA驅動寫法
    # wb = xw.Book(fullname=file_path, read_only=False, ignore_read_only_recommended=True, update_links=False)
    # wb.activate()
    # wb.api.Application.Run('refresh_dashboard')
    # wb.save()
    # wb.app.quit()
    # time.sleep(20)
    
    gc.collect() # 釋放內存記憶體

    excel_path = r"file location"
    sheet_name = "Chart"  
    range = "B3:AI63"  # chart screenshot範圍

    # chart screenshot picture
    output_path = save_excel_range_as_image(excel_path, sheet_name, range, output_directory)

    # 讀 sheet=Dashboard 
    df_dashboard = pd.read_excel(file_path, sheet_name='Dashboard')
    html_dashboard = convert_data_to_html_table(df_dashboard)

    #停機總數量 洗資料
    df_raw_data = pd.read_excel(file_path, sheet_name='raw_data', skiprows=9)
    required_columns1 = ['Group', '區域代碼', '機台號碼', '目前設備狀況', '生產狀況類別', '專案名稱', '模號', '品名', 'priority', '生產狀況敘述', '備註']
    df_raw_data_selected1 = df_raw_data[df_raw_data['目前設備狀況'] == 'STOP'][required_columns1].fillna("-")
    df_raw_data_selected1['模號'] = df_raw_data_selected1['模號'].str.replace('\n', ' ').str.strip()
    html_data1 = generate_html_from_selected_data(df_raw_data_selected1)

    # 大噸數機台開機狀況: 洗資料
    required_columns2 = required_columns1 + ['dcc週秒', '實際週秒', '週秒差異備註', 'sizemachine', 'screw', 'brands', 'tooling_tool_no']
    df_raw_data['sizemachine'] = df_raw_data['sizemachine'].apply(lambda x: str(x))
    size_machine_conditions = df_raw_data['sizemachine'].isin(['110T', '130T', '140T', '180T', '160T', '200T', '220T', '280T', '300T', '350T', '380T', '420T'])
    screw_conditions = df_raw_data['screw'].apply(lambda x: bool(re.match(r'^Ø(14|15|18|20|25|30|35)$', str(x))))
    selected_data2 = df_raw_data[(df_raw_data['目前設備狀況'] == 'STOP') & size_machine_conditions & screw_conditions][required_columns2]
    selected_data2['tooling_tool_no'] = selected_data2['tooling_tool_no'].fillna(0).astype(int).astype(str).replace('0', '-')
    selected_data2 = selected_data2.fillna("-")
    html_data2 = generate_html_from_selected_data(selected_data2)

    # make mail info -> send mail
    today = datetime.now().strftime("%Y-%m-%d")
    subject = f"Production Status - 每日開機狀態 {today}"
    recipients = [
        "outlook mail address",

]
    recipients = ";".join(recipients)
    
    # send mail
    send_email(subject, html_dashboard, output_path, html_data1, html_data2, recipients)

if __name__ == "__main__":
    main() 

