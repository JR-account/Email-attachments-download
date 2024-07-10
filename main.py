import imaplib
import email
import time
from email.parser import Parser
from email.header import decode_header
import os
import subprocess

def decode_str(s):
    value, charset = decode_header(s)[0]
    if charset:
        value = value.decode(charset)
    return value

def extract_file(filepath, extract_to):
    if filepath.endswith('.rar'):
        subprocess.call(['D:\\Program Files\\WinRAR\\WinRAR.exe', 'x', '-y', filepath, extract_to])
    elif filepath.endswith('.zip'):
        subprocess.call(['D:\\Program Files\\WinRAR\\WinRAR.exe', 'x', '-y', filepath, extract_to])

def get_att(msg, subject):
    attachment_files = []
    subject_folder = os.path.join('D:\\blockchain', subject)
    if not os.path.exists(subject_folder):
        os.makedirs(subject_folder)
    for part in msg.walk():
        file_name = part.get_filename()
        if file_name:
            h = email.header.Header(file_name)
            dh = email.header.decode_header(h)
            filename = dh[0][0]
            if dh[0][1]:
                filename = decode_str(str(filename, dh[0][1]))
            data = part.get_payload(decode=True)
            # 使用邮件主题作为文件名
            filename = f"{subject}{filename[filename.rfind('.'):]}"
            filepath = os.path.join(subject_folder, filename)
            with open(filepath, 'wb') as att_file:
                attachment_files.append(filepath)
                att_file.write(data)
            # 解压缩文件
            extract_file(filepath, subject_folder)
    return attachment_files

host = 'imap.163.com'
username = 'blockchain2024@163.com'
password = 'GMXSVFSSPOARYAF'   # 改为正确授权码

# Adding ID command
imaplib.Commands['ID'] = ('AUTH')

mail = imaplib.IMAP4_SSL(host)
mail.login(username, password)

# Upload client identity information
args = ("name", "blockchain_client", "contact", "blockchain2024@163.com", "version", "1.0.0", "vendor", "myclient")
typ, dat = mail._simple_command('ID', '("' + '" "'.join(args) + '")')
print(mail._untagged_response(typ, dat, 'ID'))

# Select inbox
status, messages = mail.select('inbox')
if status != 'OK':
    print("Failed to select inbox.")
    mail.logout()
    exit()

result, data = mail.search(None, 'ALL')
if result != 'OK':
    print("Failed to search emails.")
    mail.logout()
    exit()

mail_ids = data[0].split()

# 统计邮件数量
print(f'Messages: {len(mail_ids)}')

for i in mail_ids:
    result, data = mail.fetch(i, '(RFC822)')
    if result != 'OK':
        print(f"Failed to fetch email with ID {i}.")
        continue
    raw_email = data[0][1]
    msg = email.message_from_bytes(raw_email)

    subject = decode_str(msg.get('Subject'))
    get_att(msg, subject)

print("文件已下载完成，10秒后关闭程序！")
time.sleep(10)
mail.logout()
