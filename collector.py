import imaplib
import smtplib
import email
import os
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from io import BytesIO

# ==========================================
# НАСТРОЙКИ
# ==========================================
EMAIL_USER     = os.environ['EMAIL_USER']
EMAIL_PASS     = os.environ['EMAIL_PASS']
IMAP_SERVER    = 'imap.gmail.com'
SMTP_SERVER    = 'smtp.gmail.com'
COMMERCE_EMAIL = 'rozenbergsofi8@gmail.com'

# ==========================================
# ШАГ 1: Подключаемся к почте
# ==========================================
def connect_imap():
    print("📧 Подключаемся к почте sofiapigyukkkbgg@gmail.com ...")
    mail = imaplib.IMAP4_SSL(IMAP_SERVER)
    mail.login(EMAIL_USER, EMAIL_PASS)
    mail.select('inbox')
    print("✅ Подключились!")
    return mail

# ==========================================
# ШАГ 2: Ищем непрочитанные письма
#         с темой "Привет"
# ==========================================
def find_emails(mail):
    print("🔍 Ищем письма с темой 'Привет'...")
    
   status, messages = mail.search(None, 'UNSEEN') 
    
    email_ids = messages[0].split()
    print(f"📬 Найдено писем: {len(email_ids)}")
    return email_ids

# ==========================================
# ШАГ 3: Скачиваем Excel файлы из писем
# ==========================================
def get_attachments(mail, email_ids):
    print("📎 Скачиваем вложения...")
    
    files = []
    
    for email_id in email_ids:
        status, msg_data = mail.fetch(email_id, '(RFC822)')
        msg = email.message_from_bytes(msg_data[0][1])
        
        # Имя поставщика
        supplier_name, supplier_addr = email.utils.parseaddr(msg['From'])
        supplier = supplier_name if supplier_name else supplier_addr
        
        print(f"   📧 Письмо от: {supplier}")
        
        # Ищем Excel вложения
        for part in msg.walk():
            filename = part.get_filename()
            
            if filename and filename.endswith(('.xlsx', '.xls')):
                print(f"   📎 Нашли файл: {filename}")
                file_data = part.get_payload(decode=True)
                files.append((supplier, file_data, filename))
    
    return files

# ==========================================
# ШАГ 4: Читаем Excel и собираем таблицу
# ==========================================
def find_column(df, possible_names):
    for col in df.columns:
        for name in possible_names:
            if name.lower() in str(col).lower():
                return col
    return None

def build_summary(files):
    print("📊 Формируем сводную таблицу...")
    
    all_data = []
    
    for supplier, file_data, filename in files:
        try:
            df = pd.read_excel(BytesIO(file_data))
            
            print(f"   📋 Файл: {filename} | Строк: {len(df)}")
            print(f"   📋 Колонки: {list(df.columns)}")
            
            article_col = find_column(df, ['артикул', 'article', 'код', 'id'])
            name_col    = find_column(df, ['наименование', 'название', 'товар', 'name'])
            price_col   = find_column(df, ['цена', 'price', 'стоимость', 'cost'])
            
            if not all([article_col, name_col, price_col]):
                print(f"   ⚠️ Не нашли нужные колонки в {filename}")
                print(f"   ⚠️ Нужны: Артикул, Наименование, Цена")
                continue
            
            for _, row in df.iterrows():
                all_data.append({
                    'Поставщик':    supplier,
                    'Артикул':      row[article_col],
                    'Наименование': row[name_col],
                    'Цена':         row[price_col]
                })
                
        except Exception as e:
            print(f"   ❌ Ошибка при чтении {filename}: {e}")
    
    result = pd.DataFrame(all_data, columns=['Поставщик', 'Артикул', 'Наименование', 'Цена'])
    print(f"✅ Итого строк: {len(result)}")
    return result

# ==========================================
# ШАГ 5: Сохраняем таблицу в Excel
# ==========================================
def save_to_excel(df):
    print("💾 Сохраняем таблицу...")
    filename = 'сводная_таблица.xlsx'
    df.to_excel(filename, index=False)
    print(f"✅ Сохранено: {filename}")
    return filename

# ==========================================
# ШАГ 6: Отправляем в коммерцию
# ==========================================
def send_email(filename, total_rows):
    print(f"📤 Отправляем на {COMMERCE_EMAIL} ...")
    
    msg = MIMEMultipart()
    msg['From']    = EMAIL_USER
    msg['To']      = COMMERCE_EMAIL
    msg['Subject'] = 'Сводная таблица цен от поставщиков'
    
    body = f"""
Добрый день!

Во вложении сводная таблица цен от поставщиков.

📊 Всего позиций: {total_rows}

Таблица сформирована автоматически.
    """
    msg.attach(MIMEText(body, 'plain', 'utf-8'))
    
    with open(filename, 'rb') as f:
        attachment = MIMEBase('application', 'octet-stream')
        attachment.set_payload(f.read())
        encoders.encode_base64(attachment)
        attachment.add_header(
            'Content-Disposition',
            f'attachment; filename="сводная_таблица.xlsx"'
        )
        msg.attach(attachment)
    
    with smtplib.SMTP_SSL(SMTP_SERVER, 465) as server:
        server.login(EMAIL_USER, EMAIL_PASS)
        server.send_message(msg)
    
    print("✅ Письмо отправлено!")

# ==========================================
# ЗАПУСК
# ==========================================
def main():
    print("=" * 40)
    print("🤖 РОБОТ ЗАПУЩЕН")
    print("=" * 40)
    
    mail = connect_imap()
    email_ids = find_emails(mail)
    
    if not email_ids:
        print("📭 Писем с темой 'Привет' не найдено.")
        return
    
    files = get_attachments(mail, email_ids)
    
    if not files:
        print("📭 Excel файлов во вложениях не найдено.")
        return
    
    df = build_summary(files)
    
    if df.empty:
        print("📭 Таблица пустая.")
        return
    
    filename = save_to_excel(df)
    send_email(filename, len(df))
    
    print("=" * 40)
    print("✅ РОБОТ ЗАВЕРШИЛ РАБОТУ")
    print("=" * 40)

if __name__ == '__main__':
    main()
