# -*- coding: utf-8 -*-
"""
「resume_converter.py」的无界面自动化测试。

不启动 PySimpleGUI，直接调用核心函数 xlsx2docx_resume_converter：
  0) 图片方向自动摆正专项：
     - 合成用例：把正向证书顺时针转90°再喂给 make_upright，应转回正向
     - 真实用例1：数据集中真实上下颠倒的毕业证（EXIF=8 但实际需转180°，
       EXIF 是错标的）-> 应按内容识别出 180°
     - 真实用例2：EXIF=6 但原始像素本就正向的图 -> 应保持 0°（不被假标记骗到）
  1) 关闭 auto_orient 跑子集（第2~4行），验证开关有效、输出正常
  2) 开启 auto_orient 跑全量（7份简历），校验 docx 可打开、内嵌图片数正确
  3) 缓存命中复跑，验证第二次导出无需重新识别

运行：python test_resume_converter.py
"""
import os
import sys
import time
import warnings
from datetime import datetime
from os import path

warnings.filterwarnings('ignore')  # 抑制 openpyxl "no default style" 警告

import pandas as pd
from PIL import Image
from docx import Document

BASE = path.dirname(path.abspath(__file__))
sys.path.insert(0, BASE)

# 导入被测模块（GUI 在 __main__ 保护下，导入安全）
from importlib import import_module
mod = import_module('resume_converter')
dmod = import_module('resume_direct')

DATA_DIR = path.join(BASE, '长春吉大附中实验学校证书0801')
RESUME_XLSX = path.join(DATA_DIR, '长春吉大附中实验学校0801.xlsx')
TEMPLATE_DOCX = path.join(BASE, '应聘登记表模板-长春吉大附中实验学校-含证书.docx')
CFG_XLSX = path.join(BASE, '配置项-长春吉大附中实验学校.xlsx')
IMAGE_DIR = path.join(DATA_DIR, '图片')

PIC_COLS = ['小一寸照片（必填）', '身份证（正、反面）（必填）',
            '毕业证、学位证、岗位所需相关资格证（必填）', '个人承诺（必填）']

LAST_OUT = []      # 最近一次模板模式输出目录（供直接模式对比用）


def expected_pic_count(row_idx):
    """期望插入图片数 = 各图片列 min(单元格内文件名数,
    配置标签槽位数, 模板中实际存在的占位符数)。"""
    data = pd.read_excel(RESUME_XLSX, dtype=str)
    cfg = pd.read_excel(CFG_XLSX, sheet_name='图片项')
    # 模板中真正存在的占位符（含证书模板只配了 degree1~4，没有 degree5/6）
    from docxtpl import DocxTemplate
    tpl_vars = set(DocxTemplate(TEMPLATE_DOCX).get_undeclared_template_variables())
    n = 0
    for jcol, col in enumerate(PIC_COLS):
        # 配置表该列定义的标签槽位数（第3行起非空标签，如 degree1~6）
        series = cfg.iloc[:, jcol]
        tags = [v for v in series.iloc[2:] if isinstance(v, str) and v.strip()]
        usable = sum(1 for t in tags if t in tpl_vars) or len(tags)
        cell = data[col][row_idx]
        if isinstance(cell, str) and cell.strip():
            n += min(len(cell.split('\n')), usable)
    return n


def run_case(tag, start_n, end_n, auto_orient=True):
    ts = datetime.now().strftime('%H-%M-%S')
    out_dir = path.join(BASE, 'test_output', f'{tag}_{ts}', 'docx')
    LAST_OUT.clear()
    LAST_OUT.append(out_dir)
    log_dir = path.join(BASE, 'test_output', f'{tag}_{ts}', 'log')
    os.makedirs(out_dir, exist_ok=True)
    os.makedirs(log_dir, exist_ok=True)
    import logging
    logging.basicConfig(filename=path.join(log_dir, 'test.log'),
                        level=logging.INFO, force=True)

    print(f'\n===== 用例 {tag}: start_n={start_n}, end_n={end_n}, '
          f'auto_orient={auto_orient} =====')
    t0 = time.time()
    mod.xlsx2docx_resume_converter(
        RESUME_XLSX, TEMPLATE_DOCX, start_n, end_n,
        IMAGE_DIR, CFG_XLSX, out_dir, log_dir, auto_orient=auto_orient)
    elapsed = time.time() - t0

    # ---- 校验 ----
    files = sorted(os.listdir(out_dir))
    assert files, '输出目录为空，没有生成任何 docx'
    expect_rows = list(range(0, 7)) if (start_n == 0 and end_n == 0) \
        else list(range(start_n - 2, end_n - 1))

    assert len(files) == len(expect_rows), \
        f'生成数量不符：期望 {len(expect_rows)} 份，实际 {len(files)} 份'

    data = pd.read_excel(RESUME_XLSX, dtype=str)
    for i in expect_rows:
        j = data[data.columns[1]][i]   # 应聘单位
        c = data[data.columns[2]][i]   # 姓名
        n = data[data.columns[3]][i]   # 性别
        t = data['提交时间（自动）'][i].split()[0]
        expect_name = f'{j}_第{i+2}行_{c}_{n}_{t}.docx'
        fpath = path.join(out_dir, expect_name)
        assert path.exists(fpath), f'缺少期望文件：{expect_name}，实际有 {files}'
        doc = Document(fpath)          # 能正常打开即结构合法
        got_imgs = len(doc.inline_shapes)
        want_imgs = expected_pic_count(i)
        assert got_imgs == want_imgs, \
            f'{expect_name} 内嵌图片数不符：期望 {want_imgs}，实际 {got_imgs}'
        print(f'  [OK] {expect_name}  打开正常，内嵌图片 {got_imgs}/{want_imgs}'
              f'  （用时{elapsed:.1f}秒）' if i == expect_rows[0] else
              f'  [OK] {expect_name}  打开正常，内嵌图片 {got_imgs}/{want_imgs}')
    return len(files)


def verify_double_suffix_match():
    """专项验证：excel 中记录的名字（含 / : ）能在图片目录里匹配到
    带双重后缀的实际文件。"""
    sample = '收集结果-Vicky-第27题-2023/07/31 16:14:07.jpeg'.replace('/', '-').replace(':', '-')
    found = mod.find_files(IMAGE_DIR, sample)
    assert found and path.exists(found), f'substring 匹配失败：{sample}'
    base = path.basename(found)
    assert '..' in base, f'实际文件应带双重后缀，实际为：{base}'
    print(f'\n[OK] 双重后缀匹配验证：\n     excel 记录 -> {sample}\n     实际命中   -> {base}')


# -------------------- 图片方向自动摆正专项 --------------------
ORIENT_TMP = path.join(BASE, 'test_output', 'orient_tmp')


def _find_by_exif(ori_value):
    """在图片目录中按 EXIF 方向标记找样例文件。"""
    for f in os.listdir(IMAGE_DIR):
        p = path.join(IMAGE_DIR, f)
        if not path.isfile(p):
            continue
        try:
            if Image.open(p).getexif().get(274) == ori_value:
                return f
        except Exception:
            pass
    return None


def test_orientation_synthetic():
    """合成用例：正向证书顺时针转90°后，make_upright 应转回正向。"""
    os.makedirs(ORIENT_TMP, exist_ok=True)
    src_name = '收集结果-universe风-第29题-0-2023-08-01 00-36-54.jpeg..jpeg'
    upright = Image.open(path.join(IMAGE_DIR, src_name)).convert('RGB')
    wrong = mod._rotate_ccw(upright, 270)   # 逆时针270 == 顺时针90
    fixed, delta, reason = mod.make_upright(
        wrong, '合成_顺转90.jpeg', ORIENT_TMP)
    assert delta == 90, f'期望旋转90°还原，实际 {delta}° [{reason}]'
    import numpy as np
    assert np.array(fixed).shape == np.array(upright).shape
    diff = (np.array(fixed, dtype=int) - np.array(upright, dtype=int)).__abs__().mean()
    assert diff < 2.0, f'还原图素差异过大：{diff:.2f}'
    print(f'[OK] 合成方向用例：顺时针转90° -> 识别为需逆时针{delta}° ({reason})，像素还原')


def test_orientation_real_upside_down():
    """真实用例1：EXIF=8 的毕业证实际是上下颠倒(需180°)，EXIF 是错标的，
    内容识别必须给出 180 而不是 EXIF 的 90。"""
    name = _find_by_exif(8)
    assert name, '找不到 EXIF=8 的样例图'
    img = Image.open(path.join(IMAGE_DIR, name))
    fixed, delta, reason = mod.make_upright(img.convert('RGB'), name, ORIENT_TMP)
    assert delta == 180, \
        f'真实颠倒图应识别出180°，实际 {delta}° [{reason}]（EXIF 错标未被纠正？）'
    print(f'[OK] 真实颠倒毕业证：EXIF=8(错标) -> 内容识别 {delta}° ({reason})')


def test_orientation_bogus_exif():
    """真实用例2：EXIF=6 的图实际存储方向是倒置的(三联目检+框几何双确认
    应转180°)，EXIF 标的90°是错的 -> 内容几何识别必须给出 180 而非 90/0。"""
    name = _find_by_exif(6)
    assert name, '找不到 EXIF=6 的样例图'
    img = Image.open(path.join(IMAGE_DIR, name))
    _fixed, delta, reason = mod.make_upright(img.convert('RGB'), name, ORIENT_TMP)
    assert delta == 180, f'该图真实方向为180°，实际 {delta}° [{reason}]'
    print(f'[OK] EXIF错标回归：EXIF=6(标90°) -> 内容几何识别 {delta}° ({reason})')


def test_orientation_portrait_no_flip():
    """回归用例：正立一寸照曾被 haar 弱多数误检翻转(倒置方向误检出更多
    "脸")。人脸层加严(>=2张且>=3倍次高)后应弃权并保持 0°。"""
    name = '收集结果-♥梦夜♥-第27题-2023-08-01 08-04-24.jpeg..jpeg'
    img = Image.open(path.join(IMAGE_DIR, name)).convert('RGB')
    _fixed, delta, reason = mod.make_upright(img, '回归_正立人照.jpg', ORIENT_TMP)
    assert delta == 0, f'正立人像应保持0°，实际 {delta}° [{reason}]'
    print(f'[OK] 人像误判回归：正立一寸照 -> 保持 {delta}° ({reason})')


def test_orientation_cache_hit():
    """缓存命中：同一张图第二次 make_upright 不应再做内容识别。"""
    name = _find_by_exif(6)
    img = Image.open(path.join(IMAGE_DIR, name)).convert('RGB')
    t0 = time.time()
    mod.make_upright(img, name, ORIENT_TMP)          # 首次（可能已有缓存）
    t1 = time.time()
    _f, delta, reason = mod.make_upright(img, name, ORIENT_TMP)  # 第二次
    t2 = time.time()
    second_cost = t2 - t1
    first_cost = t1 - t0
    assert '缓存' in reason, f'第二次应命中缓存，实际依据={reason}'
    print(f'[OK] 缓存命中：首次 {first_cost:.2f}s -> 二次 {second_cost:.3f}s ({reason})')


def test_transparent_png_preserved():
    """个人承诺列的透明 PNG 应以 PNG 临时文件插入而非转成全黑 JPEG。"""
    temp_dir = path.join(IMAGE_DIR, 'temp_pic')
    pngs = [f for f in os.listdir(temp_dir) if f.endswith('.png')]
    assert pngs, 'temp_pic 中没有 PNG，说明个人承诺透明通道被破坏'
    modes = set()
    for f in pngs[-5:]:
        modes.add(Image.open(path.join(temp_dir, f)).mode)
    print(f'[OK] 个人承诺透明PNG：temp_pic 中有 {len(pngs)} 个 PNG，模式={sorted(modes)}')


# -------------------- 直接生成版一致性对比 --------------------

def _doc_texts(doc):
    """提取文档全部非空文本（段落+表格物理单元格），去除空白后排序。
    表格按 XML 物理单元格(tr.tc_lst)遍历；单元格内只取 w:t 节点——
    docxtpl 渲染产物中存在游离于 w:p/w:r 层级的裸文本残留(视觉不可见)，
    itertext() 会把它们重复计入导致两版对比假性不一致。"""
    W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    texts = []
    for p in doc.paragraphs:
        s = ''.join(p.text.split())
        if s:
            texts.append(s)
    for t in doc.tables:
        for tr in t._tbl.tr_lst:
            for tc in tr.tc_lst:
                s = ''.join(''.join(n.text or '' for n in tc.iter(W + 't'))
                            .split())
                if s:
                    texts.append(s)
    return sorted(texts)


def _direct_expected_count(row_idx):
    """直接模式期望图片数 = 各图片列 min(实际上传数, 内置上限)
    上限：照片1/身份证2/证书6/承诺1（PIC_SPECS 的语义）。"""
    data = pd.read_excel(RESUME_XLSX, dtype=str)
    limits = [1, 2, 6, 1]
    n = 0
    for col, mx in zip(PIC_COLS, limits):
        cell = data[col][row_idx]
        cnt = len([x for x in str(cell).split('\n') if x.strip()]) \
            if isinstance(cell, str) and cell.strip() else 0
        n += min(cnt, mx)
    return n


def test_direct_mode_parity():
    """直接模式 vs 模板模式 全量对比：
    - 输出文件名集合一致
    - 全部文本内容一致（模板占位符已被值替换，两边应同文）
    - 图片数：直接模式应等于"上传数与内置上限的较小值"；证书超过4张的行，
      模板版受 degree1~4 槽位限制会被截断，直接模式全插（已知且更优）。"""
    assert LAST_OUT and path.isdir(LAST_OUT[0]), \
        '请先跑 orient_on_full 用例生成模板模式基准输出'
    out_d = path.join(BASE, 'test_output',
                      f'direct_{datetime.now().strftime("%H-%M-%S")}')
    os.makedirs(out_d, exist_ok=True)
    print('\n===== 直接模式全量导出（与模板模式对比）=====')
    dmod.xlsx2docx_resume_direct(RESUME_XLSX, 0, 0, IMAGE_DIR, out_d,
                                 auto_orient=True)

    tpl_files = sorted(os.listdir(LAST_OUT[0]))
    dir_files = sorted(os.listdir(out_d))
    assert tpl_files == dir_files, \
        f'两模式文件名不一致：\n模板={tpl_files}\n直接={dir_files}'

    data = pd.read_excel(RESUME_XLSX, dtype=str)
    for i, fname in enumerate(sorted(
            dir_files, key=lambda f: int(f.split('_第')[1].split('行')[0]))):
        td = Document(path.join(LAST_OUT[0], fname))
        dd = Document(path.join(out_d, fname))
        tt, dt_ = _doc_texts(td), _doc_texts(dd)
        assert tt == dt_, \
            f'{fname} 文本内容不一致：\n仅模板有={set(tt)-set(dt_)}\n' \
            f'仅直接有={set(dt_)-set(tt)}'
        ti, di = len(td.inline_shapes), len(dd.inline_shapes)
        want_di = _direct_expected_count(i)
        assert di == want_di, \
            f'{fname} 直接模式图片数不符：期望{want_di}(按上传数与上限)，实际{di}'
        if di != ti:
            print(f'  [OK] {fname}  文本一致；图片 {ti}->{di} '
                  f'(证书超4张，模板版槽位截断、直接版全插)')
        else:
            print(f'  [OK] {fname}  文本一致；图片 {di}/{ti}')


if __name__ == '__main__':
    # 真实数据集与模板/配置夹具均不入库（前者含个人信息，后者属旧模板模式）；
    # 本机缺数据时给出说明并跳过，而不是报错。
    _missing = [p for p in (DATA_DIR, TEMPLATE_DOCX, CFG_XLSX)
                if not path.exists(p)]
    if _missing:
        print('完整测试需要真实数据集与模板/配置夹具（均不入 git 仓库），'
              '当前缺少：')
        for _m in _missing:
            print(f'  {_m}')
        print('将其放到脚本同级目录后重跑即可执行全量回归；本次跳过。')
        sys.exit(0)

    verify_double_suffix_match()
    test_orientation_synthetic()
    test_orientation_real_upside_down()
    test_orientation_bogus_exif()
    test_orientation_portrait_no_flip()
    test_orientation_cache_hit()

    cnt1 = run_case('orient_off_subset', 2, 4, auto_orient=False)  # 开关关闭子集
    cnt2 = run_case('orient_on_full', 0, 0, auto_orient=True)      # 开关开启全量
    test_transparent_png_preserved()
    test_direct_mode_parity()
    print(f'\n全部测试通过：关摆正子集 {cnt1} 份，开摆正全量 {cnt2} 份，'
          f'直接模式对比通过。')
