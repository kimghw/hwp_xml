import win32com.client
hwp = win32com.client.GetObject('HWPFrame.HwpObject')
hwp.Open(r'C:\hwp_xml\test_step5.hwp')
print("파일 열기 완료")
