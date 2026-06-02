# 证件照 Word 排版生成器

把一张或多张照片处理成标准证件照尺寸，并在 Word 文档中排成网格，**可直接打印后裁剪**使用。

默认规格为 **1 寸证件照（2.5cm × 3.5cm）**。

## 功能特性

- 自动按证件照比例**居中裁剪**（不变形），也可选择保留整张照片（留白）。
- 在 Word 中以**真实物理厘米尺寸**插入，打印即为准确尺寸。
- 网格排版 + 可选的**裁剪边线**，方便裁切。
- 支持 1 寸 / 小 1 寸 / 2 寸 / 小 2 寸 / 身份证等常见规格，也支持自定义尺寸。
- 支持 A4、5 寸相纸、6 寸相纸。
- 自动修正手机照片的拍摄方向（EXIF）。

## 安装依赖

```bash
pip install -r requirements.txt
```

## 快速开始

1. 把你的照片（如 `照片1.jpg`、`照片2.jpg`）放进一个文件夹，例如 `photos/`。
2. 运行：

```bash
python make_photo_doc.py -i photos -o 证件照.docx
```

3. 打开 `证件照.docx`，打印时选择 **"实际大小 / 100%"**，不要勾选"适应页面"缩放，
   即可得到尺寸准确的 1 寸证件照。

## 常用示例

```bash
# 单张照片填满一整页 A4（1 寸）
python make_photo_doc.py -i 照片.jpg -o 证件照.docx

# 多张照片，每张各占一整页
python make_photo_doc.py -i photos/ -o 证件照.docx

# 5 寸相纸版式，每张照片冲印 8 张
python make_photo_doc.py -i 照片.jpg --page 5r --copies 8

# 2 寸照片
python make_photo_doc.py -i 照片.jpg --size 2inch

# 自定义尺寸 30mm × 40mm，不画裁剪线
python make_photo_doc.py -i 照片.jpg --size 30x40 --no-border

# 保留整张照片（留白，不裁剪）
python make_photo_doc.py -i 照片.jpg --fit
```

## 参数说明

| 参数 | 说明 | 默认值 |
| --- | --- | --- |
| `-i, --input` | 输入图片文件或目录（可多个） | 必填 |
| `-o, --output` | 输出 Word 文件名 | `证件照.docx` |
| `--size` | 照片规格：`1inch`/`small1inch`/`2inch`/`small2inch`/`id`，或 `宽x高`（毫米） | `1inch` |
| `--page` | 纸张：`a4` / `5r` / `6r` | `a4` |
| `--copies` | 每张照片打印的份数；`0` 表示填满整页 | `0` |
| `--gap` | 照片之间的间距（毫米） | `2` |
| `--margin` | 页边距（毫米） | `8` |
| `--no-border` | 不绘制裁剪边线 | 关 |
| `--fit` | 保留整张照片（留白），而非居中裁剪 | 关 |
| `--dpi` | 重采样目标 DPI | `300` |

## 常见证件照规格参考

| 规格 | 尺寸（宽 × 高） |
| --- | --- |
| 1 寸 | 25mm × 35mm |
| 小 1 寸 | 22mm × 32mm |
| 2 寸 | 35mm × 49mm |
| 小 2 寸（护照/签证） | 35mm × 45mm |
| 二代身份证 | 26mm × 32mm |
