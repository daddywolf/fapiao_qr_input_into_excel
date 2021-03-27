# 01,04,037002000104,30814076,294.34,20210318,11910009001762141289,22C4,
import time

# 运行前先执行pip3 install pykeyboard
from pykeyboard import *

# 把需要输出的内容谢在这里
excel_info = ['fpdm', 'fphm', 'bhsje', 'kprq', 'jym']


class Fapiao(object):

    def reconize_fapiao_information(self, info):
        list = info.split(",")
        if len(list) == 9:
            print("识别成功", end=":")
        else:
            print("识别失败，请重新扫描")
            return
        if list[0] == "01" and list[1] == "04":
            print("普通发票")
        elif list[0] == "01" and list[1] == "01":
            print("专用发票")
        elif list[0] == "01" and list[1] == "10":
            print("电子发票")
        return {
            'first': list[0],
            'second': list[1],
            'fpdm': list[2],
            'fphm': list[3],
            'bhsje': list[4],
            'kprq': list[5],
            'jym': list[6],
            'other': list[7],
            'other2': list[8],
        }


if __name__ == "__main__":
    k = PyKeyboard()
    while 1:
        bar = input("请扫描Concur条码...\n")
        bar = 'FC60F6627B374FA9A1C8'
        qr = input("请扫描发票二维码...\n")
        fapiao = Fapiao()
        # fapiao.reconize_fapiao_information(qr)
        # dic = fapiao.reconize_fapiao_information("01,10,037002000311,43951776,121.58,20210327,63863934212965558966,F179,")
        dic = fapiao.reconize_fapiao_information(
            "01,04,037002000104,30814076,294.34,20210318,11910009001762141289,22C4,")
        print(dic)
        # 自动化Excel
        k.press_keys(['command', 'tab'])
        time.sleep(1)
        for i in bar:
            k.tap_key(i)
        k.tap_key('tab')
        for i in excel_info:
            for j in dic[i]:
                k.tap_key(j)
            k.tap_key('tab')
        k.tap_key('return')
        k.press_keys(['command', 'tab'])
