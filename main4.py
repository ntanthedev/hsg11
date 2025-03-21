import re
import time

import pytesseract
from PIL import Image
import requests
import io
from openpyxl import Workbook
import multiprocessing

workbook = Workbook()
sheet = workbook.active

start = 1
end = 1202

from joblib import Parallel, delayed


def crawl(sbd):
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
    while True:
        try:
            t = time.time() * 1000
            rs = requests.Session()
            rs.get(f"http://hatinh.edu.vn/tracuudiemthihsg",verify=False)
            data = io.BytesIO(rs.get(
                f"http://hatinh.edu.vn/api/Common/Captcha/getCaptcha?returnType=image&site=32982&width=150&height=50&t={t}").content)

            img = Image.open(data)
            pix = img.load()
            for y in range(img.size[1]):
                for x in range(img.size[0]):
                    img_color = pix[x, y]
                    if len([img_color for img_color in img_color if img_color > 50]):
                        pix[x, y] = (255, 255, 255)
                    else:
                        pix[x, y] = (0, 0, 0)

            capcha_text = re.sub("\\s", "", pytesseract.image_to_string(img))
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
                'keyword': str(sbd),
                'BDC_UserSpecifiedCaptchaId': capcha_text,
                'captcha_check': capcha_text,
                'captcha_code': capcha_text,
                '_t': t,
            }

            response = rs.post('http://hatinh.edu.vn/', params=params, data=data)
            if response.text != "BotDetect" and "Nhập sai mã bảo mật" not in response.text:
                data = re.findall(r"<td  >(.*?)</td>", response.text)
                print(data)
                return data
        except BaseException as e:
            print(e)

num_cores = multiprocessing.cpu_count()

output = Parallel(n_jobs=num_cores)(delayed(crawl)(i) for i in range(start, end+1))

for v in output:
    sheet.append(v)

workbook.save(f"uwu.xlsx")
