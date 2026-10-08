# -*- coding: utf-8 -*-
"""
ResumeForge 简历批量生成工具（excel 简历汇总表 -> word 简历）
由开发期 notebook 转换而来，
逻辑与原 notebook 保持一致，仅做脚本化整理。

本版本解决的问题：
    新版腾讯文档将图片改为附件下载，下载得到的文件名存在"双重后缀"
    （如 xxx.jpeg..jpeg），且 excel 单元格中记录的文件名含 / 和 :
    （如 2023/07/21 17:03:13.png），与实际文件名不一致。
    处理方式：把单元格中的 / 和 : 替换为 -，再用子串匹配的方式
    （find_files）在图片目录中查找包含该名字的实际文件。

需要安装的三方依赖包：
    docxtpl
    pandas
    PySimpleGUI (4.x, 例如 pip install PySimpleGUI==4.60.5.1)
    Pillow
    openpyxl
    numpy
    opencv-python
    rapidocr-onnxruntime

新增功能：图片方向自动摆正（auto_orient，默认开启，GUI 可勾选关闭）
    插入 word 前自动把图片转正，三层策略：
        1) 人脸检测：4 个方向中仅一个方向检出人脸则采用（适合一寸照）
        2) OCR 框几何：RapidOCR(关闭内置行级纠偏)对 4 个方向做布局分析。
           实测其对 ±90° 旋转的中文仍能整段识别(分数无法区分正侧)，
           但检测框形状不会骗人——真横排行 宽>>高。故以 conf×len 加权的
           平均框宽高比为方向证据，取最大且明显领先(>=1.2倍次优)者
           （适合证书等文档；四向均无有效文字则弃权）
        3) EXIF 兜底：内容判不出来才参考 EXIF 方向标记
           （实测腾讯文档转存图片的 EXIF 常有错标，故不盲信）
    识别结果缓存在图片目录下 `_方向缓存.json`（文件名+大小+时间戳），
    同一批图片重复导出时无需重新识别。

用法（GUI 模式，与原 notebook 一致）：
    python resume_converter.py
"""
from openpyxl import load_workbook
from docxtpl import InlineImage
from docx.shared import Mm
from docxtpl import DocxTemplate
import pandas as pd
import requests
from io import BytesIO
from os import path
from PIL import Image
import time
from datetime import datetime
import sys
import shutil
import os
import PySimpleGUI as sg
import logging
import logging.config
import threading
import traceback
import asyncio
import aiohttp
from pathlib import Path
import warnings
import json
import numpy as np
import cv2


# 用于返回指定路径下特定文件的绝对路径，返回文件的文件名中须包含给定字符串
def find_files(path, pattern):
    for root, dirs, files in os.walk(path):
        for name in files:
            if pattern in name:
                return os.path.join(root, name)
    return None


# 配置表列名与数据表列名可能差一个“（必填）”后缀（例如配置写
# “硕士毕业学校（必填）”，数据表实际列名是“硕士毕业学校”）。
# 取列时先精确匹配，再剥离后缀做包含匹配，匹配不到才报错。
def _resolve_col(df, colname):
    if colname in df.columns:
        return colname
    core = str(colname).replace('（必填）', '')
    if core:
        for c in df.columns:
            if core in str(c):
                return c
    raise KeyError(f'数据表缺少列：{colname}')


# ==================== 图片方向自动识别（新增） ====================
# 让文字或人脸朝上的三层策略，详见文件头部说明。
_T = getattr(Image, 'Transpose', Image)  # 兼容 Pillow 10 前后的常量位置
_ROT_OPS = {0: None, 90: _T.ROTATE_90, 180: _T.ROTATE_180, 270: _T.ROTATE_270}
_ORIENT_CACHE_NAME = '_方向缓存.json'
_ORIENT_RULE_VER = 4         # 判定规则版本，变更时使旧缓存失效
_ocr_model = None            # RapidOCR 单例，首次调用时才加载
MIN_BEST_SCORE = 4.0         # OCR 方向候选的最低文字得分门槛
ASPECT_MARGIN = 1.2          # 选定方向的加权框宽高比需领先次优的倍数


def _rotate_ccw(img, delta):
    """把 PIL 图片逆时针旋转 delta 度（delta 取 0/90/180/270）。"""
    op = _ROT_OPS.get(delta % 360)
    return img if not op else img.transpose(op)


def _get_ocr():
    global _ocr_model
    if _ocr_model is None:
        from rapidocr_onnxruntime import RapidOCR
        # 必须关闭内置的行级方向分类器(cls)：它会把倒置文字行自动翻转后再识别，
        # 导致上下颠倒的图片在0度方向也拿到高分，整图方向信号被抵消
        _ocr_model = RapidOCR(use_cls=False)
    return _ocr_model


def _load_orient_cache(image_path):
    cache_file = path.join(image_path, _ORIENT_CACHE_NAME)
    if path.exists(cache_file):
        try:
            with open(cache_file, 'r', encoding='utf-8') as fp:
                return json.load(fp)
        except Exception:
            return {}
    return {}


def _save_orient_cache(image_path, cache):
    cache_file = path.join(image_path, _ORIENT_CACHE_NAME)
    try:
        with open(cache_file, 'w', encoding='utf-8') as fp:
            json.dump(cache, fp, ensure_ascii=False, indent=1)
    except Exception:
        pass  # 缓存写失败不影响主流程


def _face_counts(pil_img):
    """图片按逆时针 0/90/180/270 度四个方向各检测一次人脸，返回数量列表。"""
    gray0 = cv2.cvtColor(np.array(pil_img.convert('RGB')), cv2.COLOR_RGB2GRAY)
    cascade = cv2.CascadeClassifier(
        cv2.data.haarcascades + 'haarcascade_frontalface_default.xml')
    counts = []
    for k in range(4):
        gray = np.rot90(gray0, k) if k else gray0
        faces = cascade.detectMultiScale(gray, scaleFactor=1.1,
                                         minNeighbors=5, minSize=(48, 48))
        counts.append(int(len(faces)))
    return counts


def _ocr_layout(pil_img):
    """OCR 布局分析：返回 (得分Σconf×len, 行数, 加权检测框宽高比)。

    重要实测结论：RapidOCR 对 ±90° 旋转的中文行仍有很高的识别置信度
    （竖条区域也能整段读出），因此单靠文字分数无法区分正放与侧放；
    但检测框的几何形状不会骗人——真横排行 宽>>高，旋转后的行 高>>宽。
    以 conf×len 为权重求平均框宽高比（上限截断到10）作为方向证据。
    """
    w, h = pil_img.size
    scale = 1600.0 / max(w, h)
    if scale < 1:
        pil_img = pil_img.resize((int(w * scale), int(h * scale)))
    arr = np.array(pil_img.convert('RGB'))[:, :, ::-1]  # RGB -> BGR
    result, _elapse = _get_ocr()(arr)
    total = lines = wsum = 0.0
    if result:
        lines = float(len(result))
        for box, text, conf in result:
            xs = [p[0] for p in box]
            ys = [p[1] for p in box]
            bw = max(xs) - min(xs) + 1
            bh = max(ys) - min(ys) + 1
            c = float(conf)
            L = max(len(text), 1)
            total += c * L
            wsum += c * L * min(bw / bh, 10.0)
    aspect = (wsum / total) if total > 0 else 0.0
    return total, lines, aspect


def make_upright(pil_img, cache_key, image_path, log_func=None):
    """判断并执行图片摆正。

    返回 (摆正后的PIL图, 逆时针旋转角度delta, 判断依据)。
    结果以 cache_key(一般用文件名)+大小+修改时间 为键缓存到 image_path 下。
    """
    cache = _load_orient_cache(image_path)
    try:
        st = os.stat(path.join(image_path, cache_key))
        stamp = f'{st.st_size}_{int(st.st_mtime)}'
    except OSError:
        stamp = ''
    hit = cache.get(cache_key)
    # 双方都取不到文件戳(stamp='')时也允许命中，仅要求键(文件名)一致；
    # 规则版本不一致视为失效，重新判定
    if hit and hit.get('stamp', '') == stamp \
            and hit.get('ver') == _ORIENT_RULE_VER:
        delta = int(hit.get('delta', 0))
        reason = f'缓存[{hit.get("reason", "")}]'
        if log_func:
            log_func(f'方向缓存命中：{cache_key} 旋转{delta}°')
        return _rotate_ccw(pil_img, delta), delta, reason

    exif_ori = pil_img.getexif().get(274)
    delta, reason = 0, None

    # 第1层：人脸检测（毫秒级）。haar 在证件纹理上容易误检，
    # 实测存在正立人像在倒置方向误检出更多"脸"导致被翻转的案例，
    # 因此要求明显领先才采信：最高计数>=2 且 >=3 倍次高；否则弃权交给下层
    try:
        counts = _face_counts(pil_img)
    except Exception:
        counts = [0, 0, 0, 0]
    mx = max(counts)
    second_mx = sorted(counts)[-2]
    if mx >= 2 and mx >= 3 * max(second_mx, 1):
        delta = counts.index(mx) * 90
        reason = f'人脸(counts={counts})'

    # 第2层：OCR 框几何判定。四方向各做一次布局分析，
    # 在有有效文字信号(得分>=MIN_BEST_SCORE)的方向里选加权框宽高比最大者；
    # 需明显领先(>=1.2倍次优)才采信，否则弃权。
    if reason is None:
        try:
            ocr_results = [_ocr_layout(_rotate_ccw(pil_img, k * 90))
                           for k in range(4)]
        except Exception:
            ocr_results = [(0.0, 0.0, 0.0)] * 4
        scores = [t if np.isfinite(t) else 0.0 for t, _n, _a in ocr_results]
        lines = [n if n and n > 0 else 0 for _t, n, _a in ocr_results]
        aspects = [a if np.isfinite(a) else 0.0 for _t, _n, a in ocr_results]
        valid = [k for k in range(4) if scores[k] >= MIN_BEST_SCORE]
        if valid:
            pick = max(valid, key=lambda k: (aspects[k], scores[k]))
            second_asp = sorted((aspects[k] for k in valid), reverse=True)[1] \
                if len(valid) > 1 else 0.0
            if aspects[pick] >= 1.2 * max(second_asp, 0.01):
                delta = pick * 90
                reason = f'OCR几何(aspect={[round(a, 1) for a in aspects]}, ' \
                         f'分={[round(s) for s in scores]}, 选{delta}°)'

    # 第3层：EXIF 兜底。实测腾讯文档转存图片 EXIF 常标错，故只在内容判断不出时使用
    if reason is None and exif_ori in (3, 6, 8):
        delta = {3: 180, 6: 270, 8: 90}[exif_ori]
        reason = f'EXIF兜底(orientation={exif_ori})'

    if reason is None:
        delta, reason = 0, '无法判断保持原样'

    out = _rotate_ccw(pil_img, delta)
    cache[cache_key] = {'stamp': stamp, 'delta': delta, 'reason': reason,
                        'ver': _ORIENT_RULE_VER}
    _save_orient_cache(image_path, cache)
    if log_func:
        log_func(f'方向识别：{cache_key} 旋转{delta}° [{reason}]')
    return out, delta, reason


# 参数resume_table是腾讯文档导出的excel简历汇总表, word_template是word简历模板,
# default_pic是默认空白图片, image_path是图片保存路径, cfg_table是配置excel简历列名与word简历模板中每项对应关系的表
#def xlsx2docx_resume_converter(resume_table, word_template, default_pic, image_path, cfg_table, docx_path):
# start_n, end_n, 表示要导出的起始行号和结束行号
# auto_orient, 是否在插入前自动摆正图片方向（人脸/文字朝上）
def xlsx2docx_resume_converter(resume_table, word_template, start_n, end_n, image_path, cfg_table, docx_path, log_path, auto_orient=True):
    # 程序开始时间
    start_time = time.time()
    print(f"{time.strftime('%X')}开始导出")
    data = pd.read_excel(resume_table, dtype=str)
    # 取excel表格的行数
    total = data.shape[0]
    if start_n==0 and end_n==0:
        n1 = 0
        n2 = total
    else:
        n1 = start_n - 2 if start_n >= 2 else 0
        n2 = end_n - 1 if end_n >= 2 else 0

    tpl = DocxTemplate(word_template)

    cfg_str = pd.read_excel(cfg_table, sheet_name='字符项')
    n_cfg_str = cfg_str.shape[1]
    cfg_date = pd.read_excel(cfg_table, sheet_name='日期项')
    n_cfg_date = cfg_date.shape[1]
    cfg_pic = pd.read_excel(cfg_table, sheet_name='图片项')
    # 配置表的列数
    n_cfg_pic = cfg_pic.shape[1]
    # 配置表的行数
    c_cfg_pic = cfg_pic.shape[0]
    # 为原始表增加下载图片本地路径列
    logfilename1 = path.join(log_path, '添加列前.xlsx')
    data.to_excel(logfilename1)
    data = pd.concat([data, pd.DataFrame(columns=[f"{cfg_pic.iloc[:,j].name}_localpath" for j in range(n_cfg_pic)])])
    logfilename2 = path.join(log_path, '添加列后.xlsx')
    data.to_excel(logfilename2)
    # 将nan转为空字符串''，避免显示输出的在word文档中
    g = lambda x : x if not pd.isna(x) else ''
    # 如果表中不存在下载图片的本地路径，就改为默认空白图片
    # f = lambda x : x if not pd.isna(x) else default_pic 不存在可以不设标签值，标签自动转为空白

    # 异步多协程下载图片
    #asyncio.run(coroutine_multidownload(cfg_pic, data, image_path, default_pic))
    #print(data)
    # 从带图片的excel中提取图片
    # 新的腾讯文档将图片改为附件下载，不直接放到excel中
    # extract_pictures(excel_resume_table=resume_table, pd_table=data, image_path=image_path, start_n=start_n, end_n=end_n)

    excel_filename,extension = path.splitext(path.basename(resume_table))
    logfilename3 = path.join(log_path, f'{excel_filename}_简历列表及下载图片的本地路径.xlsx')
    data.to_excel(logfilename3)

    for i in range(n1, n2):
        context = {}
        for j in range(n_cfg_str):
            key1 = cfg_str.iloc[:,j].iloc[0]
            colname = _resolve_col(data, cfg_str.iloc[:,j].name)
            value1 = g(data[colname][i])
            context[key1] = value1
            logging.info(f"简历表文字内容：{key1} {colname} {value1}")
        for j in range(n_cfg_date):
            key2 = cfg_date.iloc[:,j].iloc[0]
            colname = _resolve_col(data, cfg_date.iloc[:,j].name)
            raw2 = g(data[colname][i])
            value2 = raw2.split()[0] if raw2 else ''
            context[key2] = value2
            logging.info(f"简历表日期：{key2} {colname} {value2}")
        for j in range(n_cfg_pic):
            # 从表中取列，存为series系列
            series_t = cfg_pic.iloc[:,j]
            # 前两行是图片长度和宽度px值
            key_w = series_t.iloc[0]
            key_h = series_t.iloc[1]
            colname = _resolve_col(data, series_t.name)
            # 新建模板标签空数组
            # jinja2_tag = []
            # 简历大表中存储地址的数组
            picnames = g(data[f"{colname}"][i])
            # 该列没有上传图片：跳过不插入（模板标签留空即可）
            if not picnames:
                continue
            # 将excel中的收集结果-Sunday🇨🇳-第26题-2023/05/05 16:56:20.jpeg
            #       转为 收集结果-Sunday🇨🇳-第26题-2023-05-05 16-56-20.jpeg
            picnames = picnames.replace('/', '-')
            picnames = picnames.replace(':', '-')
            print(picnames)
            # 将换行分隔的多个文件名转为数组，像身份证正反面，学历学位证书是多个文界面写在同一个单元格中，需要进行分列处理
            pathlist = [p for p in picnames.split('\n') if p]
            print(pathlist)
            # c_cfg_pic为配置表行数
            for k in range(c_cfg_pic-2):
                keyn = series_t.iloc[k+2]
                if keyn and not pd.isna(keyn):
                    # 如果 k 大于 单元格中文件名的数量，说明配置的图片数量大于应聘者实际上传的图片数量，
                    # 以实际上传的图片数量为准，多余的配置项无图片可插，因此不执行插入命令
                    if type(pathlist)==list and k < len(pathlist):
                        # 新版腾讯文档下载的excel表中不包含图片，只包含图片文件名，需要下载附件解压后得到图片
                        #pic_abspath = path.join(image_path, pathlist[k])
                        # pathlist[k]中的文件名比实际下载的文件名少后缀，
                        # 需要从下载图片中查找文件名包含从excel读取的文件名的文件，获取其绝对路径
                        # 从excel读取的文件名是收集结果-吉盛公司-第27题-2023/07/21 17:03:13.png
                        # 实际文件名是         收集结果-吉盛公司-第27题-2023-07-21 17-03-13.png..png
                        pic_abspath = find_files(image_path, pathlist[k])
                        # 如果有包含文件名的图片，则执行插入命令，没有则不执行
                        if pic_abspath:
                            # 在存放图片的路径下，创建一个临时图片文件夹，用于存放临时图片
                            temp_folder = path.join(image_path, 'temp_pic')
                            if not path.exists(temp_folder):
                                os.makedirs(temp_folder)

                            src_img = Image.open(pic_abspath)
                            # 自动摆正图片方向（人脸/文字朝上）
                            if auto_orient:
                                src_img, ori_delta, _rs = make_upright(
                                    src_img, path.basename(pic_abspath),
                                    image_path, logging.info)
                                logging.info(f'插入前摆正：{pathlist[k]} '
                                             f'逆时针旋转{ori_delta}°')

                            # 个人承诺 前面是png透明图片，转换后变成全黑图片，
                            # 保存为PNG以保留透明通道（其余列转JPEG）
                            if j == n_cfg_pic - 1:
                                tempfilename = f'{str(int(time.time() * 1000000))}.png'
                                temp_abspath = path.join(temp_folder, tempfilename)
                                src_img.save(temp_abspath, format='PNG')
                            else:
                                # 为了解决原始图片格式不兼容，插入失败问题，非个人承诺签名图片，转换图片格式
                                tempfilename = f'{str(int(time.time() * 1000000))}.jpeg'
                                temp_abspath = path.join(temp_folder, tempfilename)
                                src_img.convert('RGB').save(temp_abspath, format='JPEG')
                            context[keyn] = InlineImage(tpl, temp_abspath, width=Mm(key_w),height=Mm(key_h))
                            logging.info(f"简历表图片的tag标签名称和列名称：{keyn} {colname}")

        tpl.render(context)
        j = g(data[_resolve_col(data, cfg_str.iloc[:,0].name)][i])
        c = g(data[_resolve_col(data, cfg_str.iloc[:,1].name)][i])
        n = g(data[_resolve_col(data, cfg_str.iloc[:,2].name)][i])
        raw_t = g(data[_resolve_col(data, cfg_date.iloc[:,0].name)][i])
        t = raw_t.split()[0] if raw_t else ''
        docx_filename = f'{j}_第{i+2}行_{c}_{n}_{t}.docx'
        #docx_filename = '简历'+str(i+1)+'_'+j+c+n+t+'.docx'
        s2 = path.join(docx_path, docx_filename)
        tpl.save(s2)
        print(f'已导出第{i+1}份简历，姓名：{n}')

    end_time = time.time()    # 程序结束时间
    run_time = end_time - start_time    # 程序的运行时间，单位为秒
    print(f"{time.strftime('%X')}导出结束，共导出{total}份简历，用时{round(run_time,2)}秒")

def _run_guarded(func, *args, **kwargs):
    """子线程入口：把异常完整打印到 stdout（GUI 消息框可见），
    避免 pythonw 下 stderr 被丢弃、线程静默死亡看起来像“卡死”。"""
    try:
        func(*args, **kwargs)
    except Exception:
        print('转换过程中出现异常，已中止：')
        print(traceback.format_exc())


# def start_thread(resume_table, word_template, default_pic, image_path, cfg_table, docx_path):
def start_thread(resume_table, word_template, start_n, end_n, image_path, cfg_table, docx_path, log_path, auto_orient=True):
    #print('开始执行')
    # 让格式转换函数在子线程中运行（异常打印到消息框，而不是静默丢弃）
    thread = threading.Thread(
        target=_run_guarded,
        args=(xlsx2docx_resume_converter, resume_table, word_template,
              start_n, end_n, image_path, cfg_table, docx_path, log_path,
              auto_orient))
    # 下面是设置守护线程：如果在程序中将子线程设置为守护线程，则该子线程会在主线程结束时自动退出
    thread.setDaemon(True)
    thread.start()  # 启动线程

if __name__ == '__main__':



    logdatetime_str = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    logpath = path.join('log', logdatetime_str)
    if not path.exists(logpath):
        os.makedirs(logpath)
    logfilename = path.join(logpath, '脚本执行日志.log')
    logging.basicConfig(filename=logfilename,level=logging.INFO)
    # 解决Workbook contains no default style
    warnings.filterwarnings('ignore')
    #2) 定义布局，确定行数
    layout=[[sg.Radio('模板模式（需word模板+配置表）','MODE',default=True,key='-mode-tpl-',enable_events=True),
             sg.Radio('直接模式（仅需原始数据表）','MODE',key='-mode-direct-',enable_events=True)],
            [sg.Text('请选择excel简历文件  注意：图片不能重叠')],
            [sg.In(key='-XLSX-'),sg.FileBrowse(button_text = '选择文件',target='-XLSX-',file_types = (('All Files','*.xlsx'),)),],
            [sg.Text('请选择word模板文件')],
            [sg.In(key='-DOCX-'),sg.FileBrowse(button_text = '选择文件',target='-DOCX-',file_types = (('All Files','*.docx'),),key='-browse-docx-'),],
            [sg.Text('请选择列名配置文件')],
            [sg.In(key='-cfgtab-'),sg.FileBrowse(button_text = '选择文件',target='-cfgtab-',file_types = (('All Files','*.xlsx'),),key='-browse-cfg-'),],
            [sg.Text('请选择下载图片路径')],
            [sg.In(key='-picPath-'),sg.FolderBrowse(button_text = '选择文件夹',target='-picPath-'),],
            [sg.Text('如从第2到第10行有数据填2和10(第1行是列名称),都填0表示导出全部')],
            [sg.Text('起始行号：'), sg.InputText("0", key='num1', enable_events=True)],
            [sg.Text('结束行号：'), sg.InputText("0", key='num2', enable_events=True)],
            [sg.Text('word简历保存路径')],
            [sg.In(key='-docxResumePath-'),sg.FolderBrowse(button_text = '选择文件夹',target='-docxResumePath-'),],
            [sg.Text('执行过程打印消息：')],
            [sg.ML(default_text='',disabled=True,size=(50,6),reroute_stdout=True)],
            [sg.Checkbox('插入前自动摆正图片方向（人脸/文字识别）', default=True, key='-orient-')],
            [sg.Button('确定', key='run'),sg.Button('取消')]]

    #3) 创建窗口
    window=sg.Window('简历excel转word工具',layout)

    # 默认起始行号和结束行号均为0，表示导出全部简历
    num1=0
    num2=0

    #4) 事件循环
    while True:
        event,values=window.read()#窗口的读取，有两个返回值(1.事件  2.值)
        #logger.info("事件：",event,"值：",values)
        logging.info(f"windows窗口事件：{event}，值：{values}")
        if event==None or event == sg.WINDOW_CLOSED or event == '取消':#窗口关闭事件
            break
        elif event == 'num1' or event == 'num2':
            try:
                num1 = int(values['num1'])
                num2 = int(values['num2'])
                if num2 >= num1 >= 0:
                    window['num1'].update(text_color='black')
                    window['num2'].update(text_color='black')
                    window['run'].update(disabled=False)
                else:
                    window['num1'].update(text_color='red')
                    window['num2'].update(text_color='red')
                    window['run'].update(disabled=True)
            except ValueError:
                window['num1'].update(text_color='red')
                window['num2'].update(text_color='red')
                window['run'].update(disabled=True)
        elif event in ('-mode-tpl-', '-mode-direct-'):
            # 直接模式不需要模板和配置表，置灰对应输入框
            direct = bool(values.get('-mode-direct-'))
            for k in ('-DOCX-', '-cfgtab-', '-browse-docx-', '-browse-cfg-'):
                window[k].update(disabled=direct)
        elif event == 'run':
            # 用于生成日期时间文件夹
            datetime_str = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            print(f'起始行号：{num1}')
            print(f'结束行号：{num2}')
            imagepath = values['-picPath-'] # 通过附件形式下载的图片存放的路径
            print(f'存放图片的目录为“{path.abspath(imagepath)}”')
            logging.info(f'存放图片的目录为“{path.abspath(imagepath)}”')

            # 用于存放输出的word文档
            docxpath = path.join(values['-docxResumePath-'], datetime_str)
            if not path.exists(docxpath):
                os.makedirs(docxpath)
            print(f'存放word简历的目录为“{path.abspath(docxpath)}”')
            logging.info(f'存放word简历的目录为“{path.abspath(docxpath)}”')
            #start_thread(values['-XLSX-'], values['-DOCX-'], values['-defaultpic-'], values['-imagepath-'], values['-cfgtab-'], values['-docxResumePath-'])
            orient = bool(values.get('-orient-', True))
            if values.get('-mode-direct-'):
                # 直接模式：仅凭原始数据表内置映射生成，无需模板/配置表
                from importlib import import_module
                dmod = import_module('resume_direct')
                dmod.start_thread_direct(values['-XLSX-'], num1, num2,
                                         imagepath, docxpath, orient)
            else:
                start_thread(values['-XLSX-'], values['-DOCX-'], num1, num2,
                             imagepath, values['-cfgtab-'], docxpath,
                             logpath, orient)

    #5) 关闭窗口
    window.close()
    logging.info("窗口已关闭")
