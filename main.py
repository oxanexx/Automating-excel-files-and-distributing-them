import datetime
import os
import shutil
from pathlib import Path
import pandas as pd
import win32com.client as win32

## Установить формат даты
today_string = datetime.datetime.today().strftime('%m%d%Y_%I%p')
today_string2 = datetime.datetime.today().strftime('%b %d, %Y')

## Установка имен папок для вложений и архивирования
attachment_path = Path.cwd() / 'data' / 'attachments'
archive_dir = Path.cwd() / 'archive'
src_file = Path.cwd() / 'data' / 'customers.xlsx'

df = pd.read_excel(src_file)
df.head()

customer_group = df.groupby('CUSTOMER_ID')

for ID, group_df in customer_group:
    print(ID)

## Запишите каждый идентификатор, группу в отдельные файлы Excel и используйте идентификатор,
## чтобы назвать каждый файл с сегодняшней датой
attachments = []
for ID, group_df in customer_group:
    attachment = attachment_path / f'{ID}_{today_string}.xlsx'
    group_df.to_excel(attachment, index=False)
    attachments.append((ID, str(attachment)))

df2 = pd.DataFrame(attachments, columns=['CUSTOMER_ID', 'FILE'])

email_merge = pd.merge(df, df2, how='left')
combined = email_merge[ ['CUSTOMER_ID', 'EMAIL', 'FILE'] ].drop_duplicates()

# Отправка индивидуальных отчетов по электронной почте соответствующим получателям
class EmailsSender:
    def __init__(self):
        self.outlook = win32.Dispatch('outlook.application')

    def send_email(self, to_email_address, attachment_path):
        mail = self.outlook.CreateItem(0)
        mail.To = to_email_address
        mail.Subject = today_string2 + ' Report'
        mail.Body = """Please find today's report attached."""
        mail.Attachments.Add(Source=attachment_path)
        # Показать электронную почту
        mail.Display(True)
        # Отправка
        mail.Send()

email_sender = EmailsSender()
for index, row in combined.iterrows():
    email_sender.send_email(row['EMAIL'], row['FILE'])

# Переместить файлы в архив
for f in attachments:
   shutil.move(f[1], archive_dir)