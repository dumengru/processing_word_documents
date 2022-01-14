# 设置所有 word 格式
# 1. 该程序只能处理 docx 文档
# 2. word程序顶部添加所有标题样式
# 3. word内容全部按正文书写
# 4. 标题书写格式 $一级标题$...
# 5. 图片和表格书写格式: $图片$ $表格$

from docx.enum.text import WD_PARAGRAPH_ALIGNMENT


WORD_SETTING = {
    # 1. 设置正文格式
    "英文字体": "Arial",
    "中文字体": "宋体",
    "正文字号": 12,            # 磅
    "正文加粗": False,         # 默认不加粗
    "正文对齐": -1,            # -1 左对齐; 0 居中对齐; 1 右对齐

    # 2. 设置表格样式(参考: https://blog.csdn.net/xtfge0915/article/details/83480120)
    "表格样式": "Medium Grid 1 Accent 5",

    # 3. 设置图片
    "图片对齐": 0,
    "图片宽度": 14,  # 厘米
    "图片高度": 7  # 厘米
}

# 标题级别
HEADING_NUMBER = {
    "一级标题": 1,
    "二级标题": 2,
    "三级标题": 3,
    "四级标题": 4,
    "五级标题": 5,
    "六级标题": 6,
    "七级标题": 7,
}

# 段落对齐映射
ALIGNMENT_DICT = {
    -1: WD_PARAGRAPH_ALIGNMENT.LEFT,
    0: WD_PARAGRAPH_ALIGNMENT.CENTER,
    1: WD_PARAGRAPH_ALIGNMENT.RIGHT
}

