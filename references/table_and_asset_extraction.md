# 表格抽取 + 知识阶梯抽图

## 知识阶梯图（从讲义 docx 抽）

讲义 docx 第一页的"知识阶梯"通常是一张横向流程图，三个气泡：「前铺知识 / 本讲知识 / 后续知识」。这张图必须出现在 PPT 第 3 张（封面、目录之后）。

### 抽取方法（验证可用）

讲义的"知识阶梯"图通常是浮动 anchor（不嵌在段落 inline runs 里），按段落定位的 blip 找不到。可靠方法：**按图片字节大小定位**——讲义里 ≈200KB 的 png 就是知识阶梯（其他图片是页眉/页脚 logo 几 KB，或题目配图几十 KB）。

```python
import zipfile, os

def extract_knowledge_ladder(handout_docx_path, out_dir):
    """从讲义 docx 抽出知识阶梯图（最大那张 png）。"""
    out_dir = Path(out_dir); out_dir.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(handout_docx_path) as z:
        # 按文件大小降序排，第一个排除明显小的 logo
        media = [(n, z.getinfo(n).file_size) for n in z.namelist()
                 if n.startswith("word/media/") and n.lower().endswith(('.png', '.jpg', '.jpeg'))]
        media = [(n, s) for n, s in media if "header" not in n.lower() and "footer" not in n.lower()]
        media.sort(key=lambda x: -x[1])  # 大的在前
        if not media:
            return None
        biggest_name, biggest_size = media[0]
        out_path = out_dir / "知识阶梯.png"
        with z.open(biggest_name) as f, open(out_path, "wb") as g:
            g.write(f.read())
        return str(out_path)
```

### 嵌入到 slide_structure.json

知识阶梯页用 `标题-文字少` layout（无装饰干扰）：

```json
{
  "slide_num": 3,
  "slide_role": "knowledge_ladder",
  "source": {"type": "add_layout", "layout": "标题-文字少"},
  "placeholders": {"16": "知识阶梯"},
  "add_images": [{
    "image_path": "output/<case>/assets/知识阶梯.png",
    "left": 1.0, "top": 2.2, "width": 11.3
  }]
}
```

**踩坑**：之前试过 `目录` layout，发现该 layout 自带背景装饰图形，知识阶梯图会被装饰盖住或裁切。`标题-文字少` 干净，是正确选择。

## 教学设计稿表格抽取

教学设计 docx 通常含 5-9 个表格。

### 表格分类（决定放哪）

| 表格 | 出现位置 | 处置 |
|---|---|---|
| 基本信息（项目/内容）| 一、基本信息 | 附录·基本信息 |
| 课件内容结构（模块/页码/核心内容/难度/教学目标）| 二、教材分析 → 2.1 | 附录·课件内容结构 |
| 学情分析（认知阶段/预设理解/教学对策）| 三、学情分析 → 3.2 | 附录·学情分析 |
| 教学方法（教学环节/方法/实施）| 五、教学方法与策略 → 5.1 | 附录·教学方法 |
| **教学内容表格**（板书表、填表演示、月份递推表等）| 六、详细教学过程设计 内 | **插入到对应 P 号 slide** |

### 抽取脚本（按文档顺序）

```python
from docx import Document
from docx.oxml.ns import qn

def extract_all_tables(docx_path):
    """按 docx body 顺序遍历，返回 list of 2D-list（每行的每列文本）。"""
    d = Document(docx_path)
    tables = []
    for child in d.element.body.iter():
        if child.tag.split("}")[-1] != "tbl":
            continue
        data = []
        for tr in child.findall(qn("w:tr")):
            row = []
            for tc in tr.findall(qn("w:tc")):
                # ⚠️ 一个 cell 含多个段落，且段落内容常常重复 3 次（"项目项目项目"）
                #    原因：docx 编辑器把 cell 文本同步成 3 份段落（中文 / 中文 / 中文）
                #    解决：取第一段即可
                paras = tc.findall(qn("w:p"))
                txt = "".join(paras[0].itertext()).strip() if paras else ""
                row.append(txt)
            data.append(row)
        tables.append(data)
    return tables
```

### Cell 文本去重的根因

观察：教学设计 docx 里每个 cell 的 itertext() 会返回类似 `"项目项目项目"`、`"模块一：取棋子问题模块一：取棋子问题模块一：取棋子问题"` 这种重复 3 次的内容。

原因：docx 单元格内含 3 段相同内容（docx 标准做法：中文混编场景下编辑器会把文本同步到多个段落）。取 `paragraphs[0].text` 而不是 `cell.text`（后者会拼接所有段落）。

## 教学内容表格的常见类型 + 插入位置

按实际 case 验证的表格 → P 号映射：

### 必胜策略

| 表 | 表头 | 插入 P 号（教学设计明确）| 备注 |
|---|---|---|---|
| 苹果数 1-9 胜负表 | 苹果数 / 胜负 | P7-P9（板书表格）| 2 行 10 列，分 3 页演示更好（1-3、4-6、1-9 汇总）|

### 递推计数

| 表 | 表头 | 插入 P 号（教学设计明确）|
|---|---|---|
| 对折剪绳填表 | 对折次数/层数/刀口数/段数 | P10（在 P9-P13 对折剪绳范围内）|
| 折纸折痕填表 | 对折次数/块数/折痕数 | P14（在 P14-P15 折纸折痕范围内）|
| 兔子繁殖填表 | 月份/大兔子/小兔子/总数/追问 | P20（在 P18-P24 兔子繁殖范围内）|
| 台阶 1/2 级标数 | 台阶/走法数/计算过程 | P26（在 P25-P28 台阶 1/2 范围内）|
| 台阶 1/2/3 级标数 | 台阶/走法数 | P30（在 P29-P32 台阶 1/2/3 范围内）|

### 牛吃草

| 表 | 类型 |
|---|---|
| （无明显教学表格，全是 inline 列表）| —— |

### 分解质因数

| 表 | 类型 |
|---|---|
| （无明显教学表格，全是 inline 列表）| —— |

## render_pptx.py 的 add_tables 字段

```json
{
  "slide_num": 10,
  "source": {"type": "duplicate", "from_idx": 5},
  "placeholders": {"12": "填表分析"},
  "add_text_boxes": [
    {"text": "📋 对折次数 → 层数、刀口数、段数：",
     "left": 1.4, "top": 1.0, "width": 10.5, "height": 3.0, "font_size": 24}
  ],
  "add_tables": [
    {
      "data": [
        ["对折次数", "层数", "刀口数", "剪完段数"],
        ["0次（不对折）", "1层", "1个刀口", "2段"],
        ["对折1次", "2层", "2个刀口", "3段"],
        ["...", "...", "...", "..."],
        ["对折n次", "2^n层", "2^n个刀口", "2^n+1段"]
      ],
      "left": 1.0, "top": 4.3, "width": 11.3, "height": 3.0,
      "font_size": 14, "header_fill": "#4F81BD"
    }
  ]
}
```

字段：
- `data`：二维数组，第一行为表头（自动加粗 + 蓝底）
- `left, top, width, height`：位置和大小（英寸）
- `font_size`：表格内字号（按列数自适应：≤3 列 18，4-5 列 14，≥6 列 12）
- `header_fill`：表头背景色（hex，默认 #4F81BD）

## 文本框 + 表格同页布局（标准）

```
slide 上半（文本框）：top=1.0, height=3.0    (引导语、问题描述)
slide 下半（表格）：  top=4.3, height=3.0    (核心数据)
```

如果表格行多（>=7 行），表格区域可以加高到 height=4.8（top 同时调到 2.5 起）。

## 教师备课参考附录的标准位置

```
... 主体内容 ...
P_n   课后任务（最后一张内容页，通常拆 1-2 张）
P_n+1 附录·基本信息          ← 教学设计表 1
P_n+2 附录·课件内容结构      ← 教学设计表 2
P_n+3 附录·学情分析          ← 教学设计表 3
P_n+4 附录·教学方法          ← 教学设计表 4
P_n+5 结束页（"下节课再见！"）  ← 必须最后
```

附录 slide 用 `add_layout: "标题-文字少"`，标题 ph[16] 写 `"教师备课参考 · XXX"`。

### 附录布局示例

```json
{
  "slide_num": 45,
  "slide_role": "appendix_info",
  "source": {"type": "add_layout", "layout": "标题-文字少"},
  "placeholders": {"16": "教师备课参考 · 基本信息"},
  "add_tables": [{
    "data": [
      ["项目", "内容"],
      ["课程名称", "..."],
      ["适用年级", "四年级（A+班）"],
      ["课时安排", "2 小时（120 分钟）"],
      ...
    ],
    "left": 2.0, "top": 2.0, "width": 9.3, "height": 4.5,
    "font_size": 16, "header_fill": "#4F81BD"
  }]
}
```
