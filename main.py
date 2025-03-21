import re
import time

import pytesseract
from PIL import Image
import requests
import io
from openpyxl import Workbook

import urllib.parse


workbook = Workbook()
sheet = workbook.active

start = 1
end = 3
    
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
for sbd in range(start,end):
    while True:
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

        # print("Capcha is: ", (capcha_text))

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
            'gridModuleParentId': '17',
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
        response.raise_for_status()
        # print(response.text)
        assert(response.text.strip() != "")
        if response.text != "BotDetect" and "Nhập sai mã bảo mật" not in response.text:
            data = re.findall(r"<td  >(.*?)</td>", response.text)
            print(data)
            break

    sheet.append(data)
    workbook.save(f"{start} - {end - 1}.xlsx")
