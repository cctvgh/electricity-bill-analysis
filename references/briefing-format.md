
# 防溺水监控电表简报（Word公文）生成规范

> 2026-09-25/26 实测定稿（经4轮排版迭代+逐页目检验收）。简报为月度对外公文，标题格式固定为《YYYY年MM月异常（大于50元）防溺水监控电表简报》。

## 简报结构（固定）

1. **标题**：居中、黑体、黑色（约二号21.5pt）
2. **导语**：共N只（**量词必须用"只"**）、当月应收电费合计X元、较上月增减额；最高电费（地址+金额）、最低电费（地址+金额）。地址本身含括号时外层不再套括号（写"最高电费为XX地址1183元"），避免双圆括号嵌套
3. **一、重点高耗能点位**：电费最高3个点位，金额+合计+占总电费百分比，"需重点关注"
4. **二、异常增长及突发点位**：增长率最高点位点名+增长率数值；突发突增点位列举；"需排查是否存在设备漏电或被盗电情况"；地址相近多只同步异常时建议片区联动排查
5. **三、持续异常点位**：上月已列异常、本月持续走高的点位，"急需彻底处理"
6. **四、工作建议**：三条式（用户定稿版）：
   - 1、实时监控，快速响应。加强电表在线监测，发现电量异常第一时间派单排障。
   - 2、重点核查，闭环处置。对多月异常、电费突增点位及时现场核查、处置并做好记录，跟踪整改到位。
   - 3、落实责任，保障运行。压实网管责任，确保防溺水监控设备安全稳定运行。
7. **附表**：异常电表明细（用电地址、本月、上月、差值、增长率、是否上月异常）+合计行
8. **脚注**：数据来源+统计口径（当月应收电费大于50元且电表用途为"防溺水监控"）；增长率口径说明（基准增长率=相对近3个月基准期平均电费的增长率）
9. **落款**：XX分公司运维部 + 日期，右对齐右空4字，上方留盖章空间

## 公文格式参数（实测）

| 项 | 参数 |
|----|------|
| 标题 | 黑体、黑色、居中、约21.5pt |
| 正文 | 仿宋、三号16pt、首行缩进2字符、行距1.5倍 |
| 小标题（一、二、三、四） | 黑体、三号、黑色 |
| 表格 | 五号10.5pt、单元格行距单倍、地址列约6.8cm其余均分、表头跨页重复、边框深灰 |
| 页码 | "— X —" 居中页脚 |
| 页边距 | 上3.7 / 下3.5 / 左2.8 / 右2.6 cm |
| 落款 | 右对齐、右缩进约2.2cm（右空4字）、单位行与日期行上下排列右侧对齐 |

## 生成方案选择（重要教训）

**表格多的公文简报必须用 python-docx 直写，不要走 md_to_js**：
- md_to_js 的表格单元格继承全文1.5倍行距，长地址列换行后行高膨胀——实测12行明细表跨5页、全文8页；
- 反复修补 JS（紧凑行距/列宽/表头重复）4轮仍难收敛；
- python-docx 直写一次成型（同样内容4页）。

python-docx 直写要点：

```python
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

doc = Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(21), Cm(29.7)
sec.top_margin, sec.bottom_margin = Cm(3.7), Cm(3.5)
sec.left_margin, sec.right_margin = Cm(2.8), Cm(2.6)

style = doc.styles['Normal']
style.font.name = '仿宋'; style.font.size = Pt(16)
style.element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋')   # 中文字体必须同时设 eastAsia
style.paragraph_format.line_spacing = 1.5

# 标题：居中黑体黑色
t = doc.add_paragraph(); t.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = t.add_run('……简报标题……')
r.font.name = '黑体'; r.font.size = Pt(21.5); r.font.color.rgb = RGBColor(0, 0, 0)
r.element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')

# 表格：紧凑行距 + 列宽 + 表头跨页重复
tb = doc.add_table(rows=1, cols=6)
trPr = tb.rows[0]._tr.get_or_add_trPr()          # 表头重复
tblHeader = OxmlElement('w:tblHeader'); tblHeader.set(qn('w:val'), 'true')
trPr.append(tblHeader)
# 单元格段落 line_spacing = 1.0、字号 Pt(10.5)；地址列 Cm(6.8)，数值列均分

# 落款：右对齐 + 右空4字 + 盖章留白
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
p.paragraph_format.right_indent = Cm(2.2)
```

## 排版目检闭环（用户对交付质量要求高，必须逐页目检）

1. **docx → PDF**：WPS COM，**PowerShell 原生 COM 无需 pywin32**（本机未装 pywin32实测可用）：
   ```powershell
   Get-Process wps -ErrorAction SilentlyContinue | Stop-Process -Force; Start-Sleep 2
   $wps = New-Object -ComObject KWPS.Application; $wps.Visible = $false
   $doc = $wps.Documents.Open($docx, $false, $true); Start-Sleep 2
   $doc.ExportAsFixedFormat($pdf, 17); $doc.Close($false); $wps.Quit()
   ```
2. **PDF → PNG**：pymupdf（`fitz.open(pdf)` + `page.get_pixmap(dpi=130).save(...)`），不依赖 Poppler
3. **逐页 image_understanding 目检**：标题居中黑色黑体、表格无挤压断词/无竖排、落款两行同页右对齐、无 `nan%` 等异常字符（NaN 增长率显示为"—"）
4. **常见问题速查**：
   - 落款跨页（单位行在第2页底、日期被挤到第3页）→ 单位行设 `page_break_before` 保两行同页
   - 表格数值列被挤竖排 → 加宽地址列/缩字号至五号
   - 页数异常膨胀（4页变8页）→ 查表格单元格是否继承1.5倍行距
   - 标题蓝色/左对齐 → 公文标题必须黑色居中（md_to_js 默认 heading-color 为蓝色且 H1 左对齐）