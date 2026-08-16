# -*- coding: utf-8 -*-
"""Excel COM 批量打开 rw_*.xlsx 验证(真实 Excel 严格校验)"""
import os, glob
import win32com.client

files = sorted(glob.glob(os.path.join(os.path.dirname(__file__), 'out', 'rw_*.xlsx')))
# 排除模板
files = [f for f in files if '_tpl_' not in os.path.basename(f)]
excel = win32com.client.DispatchEx('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
fail = 0
try:
    for f in files:
        name = os.path.basename(f)
        try:
            wb = excel.Workbooks.Open(f, ReadOnly=True)
            n = wb.Sheets.Count
            wb.Close(False)
            print('PASS %-22s sheets=%d' % (name, n))
        except Exception as ex:
            print('FAIL %-22s %s' % (name, ex))
            fail += 1
finally:
    excel.Quit()
print('%d/%d passed' % (len(files) - fail, len(files)))
raise SystemExit(1 if fail else 0)
