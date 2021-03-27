import time

# 运行前先执行：pip3 install pykeyboard
from pykeyboard import *

# 把需要输出的内容写在这里
excel_info = ['fpdm', 'fphm', 'bhsje', 'kprq', 'jym']


class Fapiao(object):

    def analyze_fapiao_information(self, info):
        list = info.split(",")
        if len(list) != 9:
            print("识别失败，请重新扫描")
            return
        elif list[0] == "01" and list[1] == "04":
            print("识别成功，普通发票")
        elif list[0] == "01" and list[1] == "01":
            print("识别成功，专用发票")
        elif list[0] == "01" and list[1] == "10":
            print("识别成功，电子发票")
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

    def keyboard_operation(self, bar, dict=None):
        if dict is None:  # bar可以不传，但是qr产生的dict不能不传
            return
        else:
            k = PyKeyboard()
            print(dict)
            # 切换到Excel
            k.press_keys(['command', 'tab'])
            time.sleep(1)
            for i in bar:
                k.tap_key(i)
            k.tap_key('tab')
            for i in excel_info:
                for j in dict[i]:
                    k.tap_key(j)
                k.tap_key('tab')
            k.tap_key('return')
            # 切换到Python窗口
            k.press_keys(['command', 'tab'])


if __name__ == "__main__":
    while 1:
        bar = input("请扫描Concur条码...\n")
        bar = 'FC60F6627B374FA9A1C8'  # 临时使用
        qr = input("请扫描发票二维码...\n")
        fapiao = Fapiao()
        dict = fapiao.analyze_fapiao_information(qr)
        fapiao.keyboard_operation(bar, dict)
