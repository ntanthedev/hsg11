from time import time as get_time_now
from requests import Session
from io import BytesIO
from pytesseract import pytesseract
from PIL import Image
from re import sub, findall
from multiprocessing import cpu_count
from multiprocessing.dummy import Pool
from datetime import datetime
from openpyxl import Workbook, load_workbook
import os
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

MAX_WORKER = cpu_count()
MIN_STUDENT = 1 
MAX_STUDENT = 1250
STUDENT_PER_WORKER = 10

def solve(start):
    new_workbook = Workbook()
    sheet = new_workbook.active
    end = start + STUDENT_PER_WORKER
    
    if not os.path.exists('successful_captchas'):
        os.makedirs('successful_captchas')
        
    for SBD in range(start, end):
        while True:
            session = Session()
            session.get(f"http://hatinh.edu.vn/tracuudiemthihsg", verify=False)
            time_now = get_time_now()
            response = session.get(f"http://hatinh.edu.vn/api/Common/Captcha/getCaptcha?returnType=image&site=32982&width=150&height=50&t={time_now}", verify=False)
            
            captcha_image = Image.open(BytesIO(response.content))
            
            pix = captcha_image.load()
            for y in range(captcha_image.size[1]):
                for x in range(captcha_image.size[0]):
                    image_color = pix[x, y]
                    pix[x, y] = (255, 255, 255) if any(c > 50 for c in image_color) else (0, 0, 0)

            answer = sub("\s", "", pytesseract.image_to_string(captcha_image))
            
            headers = {
                'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_10_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/39.0.2171.95 Safari/537.36',
                'Accept': '*/*',
                'Accept-Encoding': 'gzip, deflate, br, zstd',
                'Accept-Language': 'en-US,en:q=0.9',
                'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
                'Referer': 'https://hatinh.edu.vn/tracuudiemthihsg',
                'X-Requested-With': 'XMLHttpRequest',
                'Host': 'hatinh.edu.vn'
            }
            # orz ngai ngo teo quang
            # orz ngan tai
            params = [
                ('module', 'Content.Listing'),
                ('moduleId', '1017'),
                ('cmd', 'redraw'),
                ('site', '32982'),
                ('url_mode', 'rewrite'),
                ('submitFormId', '1017'),
                ('moduleId', '1017'),
                ('page', ''),
                ('site', '32982'),
            ]

            data = {
                'layout': 'Decl.DataSet.Detail.default',
                'itemsPerPage': '1000',
                'pageNo': '1',
                'service': 'Content.Decl.DataSet.Grouping.select',
                'itemId': '67da9e2ad0331b62c308e4b4',
                'gridModuleParentId': '16',
                'type': 'Decl.DataSet',
                'page': '',
                'modulePosition': '0',
                'moduleParentId': '-1',
                'orderBy': '',
                'unRegex': '',
                'keyword': str(SBD),
                'BDC_UserSpecifiedCaptchaId': answer,
                'captcha_check': answer,
                'captcha_code': answer,
                '_t': time_now,
            }
            response = session.post('https://hatinh.edu.vn/', params=params, data=data)
            response.raise_for_status()
            assert (response.text.strip() != "")
            if response.text != "BotDetect" and "Nhập sai mã bảo mật" not in response.text:
                data = findall(r"<td  >(.*?)</td>", response.text)
                sheet.append(data)
                print(f"{data}")
                captcha_image.save(f'successful_captchas/{answer}.png')
                new_workbook.save(f"{start}-{start+STUDENT_PER_WORKER}.xlsx")
                break

if __name__ == "__main__":
    pool = Pool(MAX_WORKER*16)
    pool.map(solve, [(i) for i in range(MIN_STUDENT, MAX_STUDENT, STUDENT_PER_WORKER)])
