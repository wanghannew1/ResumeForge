# ResumeForge · 简历批量生成工具

从 Excel 简历汇总表批量生成 Word 简历。
**无需 Word 模板、无需配置表格**——只凭原始数据表即可产出成品简历；
附件图片自动匹配、方向自动摆正、插入时保持原始长宽比。

## 仓库结构

| 文件 | 说明 |
|---|---|
| `resume_converter.py` | 程序主体：GUI 入口、双模式切换、图片方向识别算法（`make_upright()`） |
| `resume_direct.py` | **直接生成模式引擎**（推荐用法）：仅读原始数据表产出 word 简历 |
| `test_resume_converter.py` | 无界面自动化测试（开发验证用，不参与生成） |
| `start.bat` | Windows 双击启动器 |
| `image_orientation.md` | 图片方向识别三层策略的完整设计文档 |

调用关系：

```
start.bat（双击入口）
   └─> resume_converter.py      ← GUI、双模式切换、make_upright 方向算法
            └─ 直接模式 ──> resume_direct.py  ← 实际产出 word 简历的引擎
                              └─ import 主模块的 make_upright 做插图前摆正
test_resume_converter.py       ← 仅开发验证用，不参与生成
```

## 安装

Python 3.8+：

```bash
pip install docxtpl pandas openpyxl Pillow "PySimpleGUI==4.60.5.1" numpy opencv-python rapidocr-onnxruntime python-docx aiohttp requests
```

> PySimpleGUI 必须用 4.x（5.x 起改为商业授权且 API 有变化）；OCR 引擎 rapidocr-onnxruntime 的模型随包内置，无需额外下载。

## 使用

双击 `start.bat`，或：

```bash
python resume_converter.py
```

GUI 中选择模式与输入路径后开始批量转换，产物为逐人一份的 `.docx` 简历。

### 直接生成模式（推荐）

- **输入**：原始数据表 `.xlsx`（字符列 + 日期列 + 图片列）+ 附件图片文件夹；不需要 word 模板与配置项表格，版式由代码内置常量生成
- 列名按包含关系匹配，自动兼容「（必填）」后缀等差异
- 图片列支持多张：一寸照 1 张、身份证 2 张、证书类最多 6 张全部插入、个人承诺 1 张
- 插图按**原始长宽比等比缩放**放入槽位框内——贴长边、不超出、不拉伸变形
- 插入前自动把倒置/旋转图片转正（GUI 可关）
- 也接受 pandas DataFrame 直接调用 `xlsx2docx_resume_direct()`，便于嵌入其他流水线

### 模板渲染模式（兼容旧流程）

docxtpl + word 模板 + 配置项表格三件套渲染，保留给存量模板场景；模板与配置文件因含校方信息不入库。

## 关键特性

1. **附件文件名修复**：腾讯文档附件的双重后缀（如 `xxx.jpeg..jpeg`）、单元格记录名含 `/` 与 `:` 导致与实际文件名不一致——统一替换后按子串匹配查找真实文件
2. **图片方向自动摆正**（默认开启）：人脸检测（证件照）→ OCR 检测框几何（文档类：不信 OCR 分数，信检测框形状）→ EXIF 兜底（实测常见错标，故不盲信）；结果缓存在图片目录 `_方向缓存.json`，同批重跑零开销。详见 [image_orientation.md](image_orientation.md)
3. **透明通道保留**：个人承诺签名章 PNG 以 RGBA 原样嵌入，不糊白底
4. **方向规则版本化**（`_ORIENT_RULE_VER = 4`）：算法升级后旧缓存自动失效重建

## 测试

```bash
python test_resume_converter.py
```

覆盖：合成/真实倒置样张的方向识别专项、关闭摆正的开关对照子集、开摆正全量回归、缓存命中复跑、透明 PNG 校验、直接模式与模板模式的文本一致性对比。
真实数据集与模板/配置夹具均不入仓库——克隆环境运行会提示缺少文件并跳过，属预期行为。

## 数据隐私

本工具处理真实个人信息。所有真实数据（原始表格、附件图片、生成的简历、日志）一律不入 git 仓库，`.gitignore` 已按此原则配置；请使用者同样遵守所在地区的个人信息保护法规。
