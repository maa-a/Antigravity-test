#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""recalc(LibreOffice)後にフォント名が置換されるため、xl/styles.xml を直接書き換えて
Meiryo UI に戻す。openpyxl で開き直さないのでキャッシュ済みの計算結果は保持される。"""
import re
import shutil
import sys
import zipfile

TARGET = 'Meiryo UI'
SUBS = ['WenQuanYi Zen Hei', 'Calibri', 'Arial', 'DejaVu Sans', 'Liberation Sans', 'Noto Sans CJK JP']


def fix(path):
    tmp = path + '.tmp'
    zin = zipfile.ZipFile(path)
    n = 0
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename in ('xl/styles.xml', 'xl/theme/theme1.xml'):
                txt = data.decode('utf-8')
                for s in SUBS:
                    txt, k = re.subn(r'(<(?:name|latin typeface|ea typeface|cs typeface)[^>]*?"?)%s(")'
                                     % re.escape(s), r'\g<1>%s\g<2>' % TARGET, txt)
                    n += k
                    txt, k = re.subn(r'val="%s"' % re.escape(s), 'val="%s"' % TARGET, txt)
                    n += k
                    txt, k = re.subn(r'typeface="%s"' % re.escape(s), 'typeface="%s"' % TARGET, txt)
                    n += k
                data = txt.encode('utf-8')
            zout.writestr(item, data)
    zin.close()
    shutil.move(tmp, path)
    return n


if __name__ == '__main__':
    for p in sys.argv[1:]:
        print(p, 'replaced:', fix(p))
