from NaverNewsCrawler import NaverNewsCrawler

# 사용자로 부터 기사 수집을 원하는 키워드를 입력 받음.
Input_Keyword = input('수집을 원하는 기사 키워드를 입력해주세요: ')
crawler = NaverNewsCrawler(Input_Keyword)

# 수집한 데이터를 저장할 엑셀 파일명을 입력 받음.
Input_Filename = input('수집한 데이터를 저장할 엑셀 파일명(.xlsx)을 입력해주세요: ')
crawler.get_news(Input_Filename)

# 아래코드를 실행해 이메일 발송 기능에 필요한 모듈을 임포트함.
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import smtplib
import re

# gmail 발송 기능에 필요한 계정 정보를 아래에 입력.
SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 465
## 제출하기 위해 이메일 계정의 내용은 제거함. 사용 시, 내용 입력 후 실행.
SMTP_USER = ''
SMTP_PASSWORD = ''

# 아래 코드는 메일 발송에 필요한 send_mail 함수.
def send_mail(name, addr, subject, contents, attachment=None):
    if not re.match('(^[a-zA-Z0-9_.-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$)', addr):
        print('Wrong email')
        return

    msg = MIMEMultipart('alternative')
    if attachment:
        msg = MIMEMultipart('mixed')

    msg['From'] = SMTP_USER
    msg['To'] = addr
    msg['Subject'] = name + '님, ' + subject

    text = MIMEText(contents, _charset='utf-8')
    msg.attach(text)

    if attachment:
        from email.mime.base import MIMEBase
        from email import encoders

        file_data = MIMEBase('application', 'octect-stream')
        file_data.set_payload(open(attachment, 'rb').read())
        encoders.encode_base64(file_data)

        import os
        filename = os.path.basename(attachment)
        file_data.add_header('Content-Disposition', 'attachment; filename="' + filename + '"')
        msg.attach(file_data)

    smtp = smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT)
    smtp.login(SMTP_USER, SMTP_PASSWORD)
    smtp.sendmail(SMTP_USER, addr, msg.as_string())
    smtp.close()

# 프로젝트 폴더에 있는 email_list.xlsx 파일에 이메일을 받을 사람들의 데이터가 저장되어 있음.

# 엑셀 파일의 정보를 읽어올 수 있는 모듈을 import하고, email_list.xlsx을 읽어옴.
from openpyxl import load_workbook

wb = load_workbook('email_list.xlsx')
data = wb.active

subject = '%s에 대한 뉴스 수집 자동화 메일입니다.' %Input_Keyword
contents = '이 메일은 자동화로 보내지는 메일입니다.'

# 변수 cnt를 통해 0이면 첫 번째 행. 아니면 첫 번째 행이 아님을 확인.
cnt = 0
for row in data:
    # 셀의 title이 적힌 첫 번째 행은 건너뛰기 위한 if문을 작성함.
    if cnt == 0: 
        cnt+=1
        continue

    # 행의 각 셀을 num, name, emil 변수에 할당.
    (num, name, email) = row
    send_mail(name.value, email.value, subject, contents, Input_Filename)

