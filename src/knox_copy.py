import os
import sys
import win32clipboard
import logging

from urllib import request
from PIL import Image

def copy_clipboard(data):
    HTML_ID = win32clipboard.RegisterClipboardFormat("HTML Format")
    CSV_ID = win32clipboard.RegisterClipboardFormat("Csv")

    win32clipboard.OpenClipboard(0)
    win32clipboard.EmptyClipboard()

    win32clipboard.SetClipboardData(HTML_ID, data)
    win32clipboard.SetClipboardData(CSV_ID, b'csv dump')

    win32clipboard.CloseClipboard()

def gen_template(image_url: str):
    MAX_SIZE = 3000
    NULL = b'\x00'

    width, height = 200, 200
    try:
        opener = request.build_opener()
        http_proxy = os.getenv('HTTP_PROXY')
        https_proxy = os.getenv('HTTPS_PROXY')
        if http_proxy and https_proxy:
            proxy_handler = request.ProxyHandler({
                'http': http_proxy,
                'https': https_proxy,
            })
            opener = request.build_opener(proxy_handler)

        real_width, real_height = Image.open(opener.open(image_url)).size
        height = int(real_height * width / real_width)
        print("Orig image size : ",  (real_width, real_height))

    except Exception:
        logging.exception(Exception)
        
    template = (
        b'Version:0.9      \r\n'
        b'StartHTML:000000185      \r\n'
        b'EndHTML:000003000      \r\n'
        b'StartFragment:[start]      \r\n'
        b'EndFragment:[end]      \r\n'
        b'StartSelection:[start]      \r\n'
        b'EndSelection:[end]\r\n'
        b'<!DOCTYPE HTML  PUBLIC "-//W3C//DTD HTML 4.0  Transitional//EN">\r\n'
        b'<html xmlns:v=\'urn:schemas-microsoft-com:vml\' xmlns:o=\'urn:schemas-microsoft-com:office:office\' xmlns:x=\'urn:schemas-microsoft-com:office:excel\' xmlns=\'http://www.w3.org/TR/REC-html40\'>\n'
        b'<head>\n'
        b'<meta http-equiv=\'Content-Type\' content=\'text/html; charset=utf-8\'>\n'
        b'<meta name=\'auth\' content=\'ChickenChickenChickenChicken\'>\n'
        b'<style>\n'
        b'.bgbg\n'
        b'\t{background-image:url(\'' + str.encode(image_url) + b'\');\n'
        b'\tbackground-size: cover;\n'
        b'\tbackground-repeat:no-repeat;\n'
        b'\tbackground-position:center middle;}\n'
        b'-->\n'
        b'</style>\n'
        b'</head>\n'
        b'<body><!--StartFragment-->\n'
        b'<table width=\'[width]\' height=\'[height]\'><tr>\n'
        b'<td width=\'[width]\' height=\'[height]\' class=\'bgbg\'></td> </tr>\n'
        b'</table>\n'
        b'<!--EndFragment--></body>\n'
        b'</html>'
    )

    start_fragment = template.find(b'\n<table')
    end_fragment = template.find(b'<!--EndFragment-->')

    template = (
        template
        .replace(b'[width]', str.encode(str(width)))
        .replace(b'[height]', str.encode(str(height)))
        .replace(b'[start]', str.encode(str(start_fragment).zfill(9)))
        .replace(b'[end]', str.encode(str(end_fragment).zfill(9)))
    )

    while len(template) < MAX_SIZE:
        template += NULL

    return template

if __name__ == '__main__':
    os.getenv
    copy_clipboard(gen_template(sys.argv[1]))
