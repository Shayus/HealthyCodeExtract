# -*- coding: utf-8 -*-
import os
import re
import tkinter.messagebox
import tkinter.ttk
from datetime import datetime
from tkinter import *
from tkinter import filedialog

import cv2
import numpy as np
import thulac
import xlsxwriter
from paddleocr import PaddleOCR


def Ocr_baidu(ocr, str):
    result = ocr.ocr(str, cls=True)
    data = ''
    for item in result:
        data = data + item[1][0] + '\n'
    return result, data


class MyFrm(Frame):
    def __init__(self, master):
        self.root = master
        self.screen_width = self.root.winfo_screenwidth()
        self.screen_height = self.root.winfo_screenheight()
        self.root.withdraw()  # 暂时不显示窗口来移动位置
        self.root.title('核酸检测结果识别')
        self.root.resizable(False, False)
        self.root.geometry('%dx%d+%d+%d' %
                           (540, 250, (self.screen_width - 540) / 2,
                            (self.screen_height - 360) / 2))
        self.root.deiconify()


root = Tk()
MyFrm(root)


def sel_savepath():
    save_path = filedialog.askdirectory(title='选择文件存放的位置！', initialdir=r'C:')
    lable4.configure(text=save_path)  # 重新设置标签文本


def getTime():
    d = datetime.today()
    res = datetime.strftime(d, '%Y-%m-%d %H:%M:%S')
    return d, res


def get_strtime(text):
    text = text.replace("年", "").replace("月", "").replace("日", "").replace("/", "").replace("-", "").strip()
    text = text.replace("：", "").replace(":", "").replace(" ", "").strip()
    text = re.sub("\s+", " ", text)
    t = ""
    regex_list = [
        "(\d{14})",
        "(\d{12})",
    ]
    for regex in regex_list:
        t = re.search(regex, text)
        if t:
            t = t.group(1)
            return t
    else:
        return ''


def get_img_file(file_name):
    imagelist = []
    for parent, dirnames, filenames in os.walk(file_name):
        for filename in filenames:
            if filename.lower().endswith(('.bmp', '.png', '.jpg', '.jpeg')):
                path = os.path.join(parent, filename).replace("\\", "/")
                imagelist.append(path)
        return imagelist


def getTimeDis(dtime):
    day, hour, min = 24 * 60 * 60, 60 * 60, 60
    days, hours, mins, secs = 0, 0, 0, 0
    if dtime >= day:
        days, dtime = divmod(dtime, day)
    if dtime >= hour:
        hours, dtime = divmod(dtime, hour)
    if dtime >= min:
        mins, dtime = divmod(dtime, min)
    secs = dtime
    return "{}天{}时{}分{}秒".format(int(days), int(hours), int(mins), int(secs))


def get_hesuanresult(path):
    with open(path, 'rb') as f:
        np_arr = np.frombuffer(f.read(), dtype=np.uint8)
        img = cv2.imdecode(np_arr, cv2.IMREAD_COLOR)
    g = np.zeros((img.shape[0], img.shape[1]), dtype=img.dtype)
    r = np.zeros((img.shape[0], img.shape[1]), dtype=img.dtype)
    g[:, :] = img[:, :, 1]  # 复制 g 通道的数据
    r[:, :] = img[:, :, 2]  # 复制 r 通道的数据
    g = np.mean(g)
    r = np.mean(r)
    if g > r:
        return '阴性'
    else:
        return '阳性'


def sel_imgpath():
    dir_path = filedialog.askdirectory(title='选择批量截图所在文件夹！', initialdir=r'C:')
    lable9.configure(text=dir_path)


def process():
    today, ti = getTime()
    saveFile = lable4.cget("text").strip() + '/核酸检测结果统计' + ti.replace(':', '') + '.xlsx'
    dir_path = lable9.cget("text").strip()
    workbook = xlsxwriter.Workbook(saveFile)
    # 创建工作表
    worksheet = workbook.add_worksheet()
    worksheet.set_column('A:B', 10)
    worksheet.set_column('C:C', 20)
    worksheet.set_column('D:D', 30)
    worksheet.set_column('E:F', 20)
    worksheet.set_column('G:I', 30)
    worksheet.write(0, 0, "姓名")
    worksheet.write(0, 1, "检测结果")
    worksheet.write(0, 2, "检测时间")
    worksheet.write(0, 3, "检测机构")
    worksheet.write(0, 4, "统计时间")
    worksheet.write(0, 5, "时间差")
    worksheet.write(0, 6, "文件路径")

    ocr = PaddleOCR(use_angle_cls=True, lang='ch')
    thu = thulac.thulac(seg_only=False)
    punctuation = r"""!"#$%&'()*+,-./:;<=>?@[\]^_`{|}~、!"#$%&'()*+,-./:;<=>?@[\]^_`{|}~“”？，！【】（）、。：；’‘……￥·"""
    dicts = {i: '' for i in punctuation}
    punc_table = str.maketrans(dicts)

    # 定义一个红色+黑体的格式.
    bold = workbook.add_format({'bold': 1, "color": "red"})
    files = get_img_file(dir_path)
    for i in range(len(files)):
        result, data = Ocr_baidu(ocr, files[i])
        today, ti = getTime()
        re_name = ""
        re_hesuanresult = ""
        re_jiancejigou = ""
        data = data.translate(punc_table)
        text = thu.cut(data, text=True)
        re_name = re.search('\n(.*)_np', text)
        if re_name:
            re_name = re_name.group(1)
        else:
            if re.search('\n(.*)\n若', data):
                temp = re.search('\n(.*)\n若', data)
                re_name = temp.group(1)
        re_hesuanresult = get_hesuanresult(files[i])
        time = get_strtime(data)  # 文本中的时间
        time = time.ljust(14, '0')
        try:
            ttime = datetime.strptime(time, '%Y%m%d%H%M%S')
            time = datetime.strftime(ttime, '%Y-%m-%d %H:%M:%S')
        except ValueError:
            ttime = datetime.strftime('202001010000', '%Y%m%d%H%M%S')
            time = datetime.strftime(ttime, '%Y-%m-%d %H:%M:%S')

        if re.search('机构(.*)\n', data):
            re_jiancejigou = re.search('机构(.*)\n', data).group(1).replace("：", "").replace(":", "")

        worksheet.write(i + 1, 0, re_name)
        worksheet.write(i + 1, 1, re_hesuanresult)
        if re_hesuanresult != '阴性':
            worksheet.write(i + 1, 1, re_hesuanresult, bold)
        worksheet.write(i + 1, 2, time)
        worksheet.write(i + 1, 3, re_jiancejigou)
        today, ti = getTime()
        worksheet.write(i + 1, 4, ti)
        worksheet.write(i + 1, 5, getTimeDis((today - ttime).total_seconds()))
        worksheet.write(i + 1, 6, files[i].split("/")[-1])
        lable7.configure(text='完成' + str(i+1) + '/' + str(len(files)))
        root.update()
    workbook.close()
    tkinter.messagebox.askokcancel('提示', '统计完成！\n' + saveFile)
    lable7.configure(text='')


lable1 = Label(root, text='请选择文件夹', font=("黑体", 18))
lable1.place(relx=0.15, rely=0.15)

btn1 = Button(root, text='选择', font=("黑体", 18), command=sel_imgpath)
btn1.place(relx=0.55, rely=0.15, relwidth=0.3, relheight=0.1)

lable2 = Label(root, text='设置保存路径', font=("黑体", 18))
lable2.place(relx=0.15, rely=0.3)

btn2 = Button(root, text='选择', font=("黑体", 18), command=sel_savepath)
btn2.place(relx=0.55, rely=0.3, relwidth=0.3, relheight=0.1)

lable3 = Label(root, text='保存路径:', font=("黑体", 12))
lable3.place(relx=0.1, rely=0.6)

lable8 = Label(root, text='图片路径:', font=("黑体", 12))
lable8.place(relx=0.1, rely=0.7)

lable9 = Label(root, text='C:', font=("黑体", 12))
lable9.place(relx=0.25, rely=0.7)

lable4 = Label(root, text='C:', font=("黑体", 12))
lable4.place(relx=0.25, rely=0.6)

lable5 = Label(root, text='信息学院 学工办', font=("黑体", 10))
lable5.place(relx=0.75, rely=0.8)

lable6 = Label(root, text='技术支持：10-511', font=("黑体", 10))
lable6.place(relx=0.75, rely=0.9)

lable7 = Label(root, text='', font=("黑体", 12))
lable7.place(relx=0.1, rely=0.9)

btn3 = Button(root, text='开始', font=("黑体", 18), command=process)
btn3.place(relx=0.55, rely=0.45, relwidth=0.3, relheight=0.1)

root.mainloop()
