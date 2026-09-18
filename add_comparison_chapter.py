"""
Bổ sung Chương 6: So sánh tổng quan 6 thuật toán
CURE, K-Means, K-Medoids (Clustering) | Apriori, FP-Growth, IT-Tree (Association Rule Mining)
Tất cả công thức dùng Word Equation chuẩn OMML
"""
import sys, os, re
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from lxml import etree

if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

# ─── mở file hiện có ───────────────────────────────────────────────────────
INPUT_FILE  = r'e:\KTDL\Bao_Cao_Tieu_Luan_CURE.docx'
OUTPUT_FILE = r'e:\KTDL\Bao_Cao_Tieu_Luan_CURE.docx'
doc = Document(INPUT_FILE)

# ─── tiện ích ──────────────────────────────────────────────────────────────
def set_cell_background(cell, fill_hex):
    tcPr = cell._element.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_table_borders(table, color="B0B0B0"):
    tblPr = table._element.xpath('w:tblPr')
    if tblPr:
        borders = parse_xml(
            f'<w:tblBorders {nsdecls("w")}>'
            f'<w:top    w:val="single" w:sz="6" w:space="0" w:color="{color}"/>'
            f'<w:bottom w:val="single" w:sz="6" w:space="0" w:color="{color}"/>'
            f'<w:left   w:val="none"/>'
            f'<w:right  w:val="none"/>'
            f'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="{color}"/>'
            f'<w:insideV w:val="single" w:sz="4" w:space="0" w:color="{color}"/>'
            f'</w:tblBorders>'
        )
        tblPr[0].append(borders)

def clean_xml(t):
    return t.replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')

# ─── OMML helpers ──────────────────────────────────────────────────────────
MATH_NS  = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
WORD_NS  = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

SYM_MAP = [
    (r'\neq','≠'),(r'\le','≤'),(r'\ge','≥'),(r'\in','∈'),(r'\notin','∉'),
    (r'\approx','≈'),(r'\alpha','α'),(r'\beta','β'),(r'\gamma','γ'),
    (r'\delta','δ'),(r'\epsilon','ε'),(r'\varepsilon','ε'),(r'\sigma','σ'),
    (r'\mu','μ'),(r'\pi','π'),(r'\theta','θ'),(r'\lambda','λ'),
    (r'\times','×'),(r'\cdot','·'),(r'\sqrt','√'),(r'\infty','∞'),
    (r'\sum','∑'),(r'\rightarrow','→'),(r'\Rightarrow','⇒'),(r'\forall','∀'),
    (r'\exists','∃'),(r'\cup','∪'),(r'\cap','∩'),(r'\subseteq','⊆'),
    (r'\\{','{'),(r'\\}','}'),(r'\emptyset','∅'),
]

def math_to_omml_xml(text: str) -> str:
    for latex, uni in SYM_MAP:
        text = text.replace(latex, uni)

    # tokenise: find sub/sup patterns, then bare runs
    tokens = []
    i = 0
    # pattern: base_sub^sup  | base_sub | base^sup
    pat = re.compile(
        r'([A-Za-zα-ωΑ-Ω∑∈≠≤≥×·√]+|[0-9]+(?:\.[0-9]*)?)'  # base
        r'(?:_\{([^}]*)\}|_([A-Za-z0-9α-ω]+))?'             # optional sub {…} or single-char
        r'(?:\^\{([^}]*)\}|\^([A-Za-z0-9+\-α-ω]+))?'        # optional sup {…} or single-char
    )
    pos = 0
    parts = []
    while pos < len(text):
        m = pat.search(text, pos)
        if not m or m.start() > pos + len(text):
            # plain text remainder
            seg = text[pos:]
            if seg:
                parts.append(('plain', seg))
            break
        if m.start() > pos:
            parts.append(('plain', text[pos:m.start()]))
        base = m.group(1)
        sub  = m.group(2) if m.group(2) is not None else m.group(3)
        sup  = m.group(4) if m.group(4) is not None else m.group(5)
        if sub and sup:
            parts.append(('subsup', base, sub, sup))
        elif sub:
            parts.append(('sub', base, sub))
        elif sup:
            parts.append(('sup', base, sup))
        else:
            parts.append(('plain', base))
        pos = m.end()
        if pos == m.start():   # safety: avoid infinite loop
            parts.append(('plain', text[pos]))
            pos += 1

    def r_elem(txt):
        return f'<m:r><m:t xml:space="preserve">{clean_xml(txt)}</m:t></m:r>'

    inner = ''
    for p in parts:
        kind = p[0]
        if kind == 'plain':
            inner += r_elem(p[1])
        elif kind == 'sub':
            inner += f'<m:sSub><m:sSubPr/><m:e>{r_elem(p[1])}</m:e><m:sub>{r_elem(p[2])}</m:sub></m:sSub>'
        elif kind == 'sup':
            inner += f'<m:sSup><m:sSupPr/><m:e>{r_elem(p[1])}</m:e><m:sup>{r_elem(p[2])}</m:sup></m:sSup>'
        elif kind == 'subsup':
            inner += (f'<m:sSubSup><m:sSubSupPr/><m:e>{r_elem(p[1])}</m:e>'
                      f'<m:sub>{r_elem(p[2])}</m:sub><m:sup>{r_elem(p[3])}</m:sup></m:sSubSup>')
    return f'<m:oMath xmlns:m="{MATH_NS}">{inner}</m:oMath>'

def add_inline_math(paragraph, math_text: str):
    xml_str = math_to_omml_xml(math_text)
    elem = parse_xml(xml_str)
    paragraph._p.append(elem)

def add_math_paragraph(container, text: str, style_name='Normal', bold_label=False):
    p = container.add_paragraph(style=style_name)
    p.paragraph_format.line_spacing = 1.3
    p.paragraph_format.space_after  = Pt(4)

    # split on $...$ 
    parts = re.split(r'\$([^$]+)\$', text)
    for idx, part in enumerate(parts):
        if idx % 2 == 0:
            # plain text – handle **bold**
            bold_parts = re.split(r'\*\*(.+?)\*\*', part)
            for bi, bp in enumerate(bold_parts):
                if not bp:
                    continue
                run = p.add_run()
                run.text = bp
                run.bold = (bi % 2 == 1)
                run.font.name = 'Times New Roman'
                run.font.size = Pt(12)
        else:
            add_inline_math(p, part)
    return p

def add_display_equation(container, math_text: str, eq_label: str = ''):
    p = container.add_paragraph()
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(4)
    inner = math_to_omml_xml(math_text).replace('<m:oMath ','<m:oMath xmlns:m="'+MATH_NS+'" ',1)
    # strip outer tags and rewrap in oMathPara
    inner_body = re.sub(r'^<m:oMath[^>]*>|</m:oMath>$','', inner.replace(f'xmlns:m="{MATH_NS}"',''))
    xml = (f'<m:oMathPara xmlns:m="{MATH_NS}">'
           f'<m:oMath>{inner_body}</m:oMath>'
           f'</m:oMathPara>')
    elem = parse_xml(xml)
    p._p.append(elem)
    if eq_label:
        run = p.add_run(f'   {eq_label}')
        run.font.name = 'Times New Roman'; run.font.size = Pt(11)
        run.font.color.rgb = RGBColor(100,100,100)
    return p

def add_heading_1(text: str):
    p = doc.add_heading(text, level=1)
    run = p.runs[0] if p.runs else p.add_run(text)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(14)
    run.bold = True
    run.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after  = Pt(6)
    return p

def add_heading_2(text: str):
    p = doc.add_heading(text, level=2)
    run = p.runs[0] if p.runs else p.add_run(text)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(13)
    run.bold = True
    run.font.color.rgb = RGBColor(0x2E, 0x74, 0xB5)
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after  = Pt(4)
    return p

def add_heading_3(text: str):
    p = doc.add_heading(text, level=3)
    if p.runs:
        run = p.runs[0]
    else:
        run = p.add_run(text)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(12)
    run.bold = True
    run.italic = True
    run.font.color.rgb = RGBColor(0x40, 0x40, 0x40)
    return p

BLUE_HDR  = '1F497D'
BLUE_LIGHT= 'D6E4F0'
GRAY_ROW  = 'F5F5F5'
WHITE_ROW = 'FFFFFF'
GREEN_HDR = '1E6B45'
GREEN_LT  = 'D5E8D4'

# ══════════════════════════════════════════════════════════════════
# BẮT ĐẦU CHÈN CHƯƠNG 6
# ══════════════════════════════════════════════════════════════════
doc.add_page_break()

add_heading_1('CHƯƠNG 6: SO SÁNH TỔNG QUAN 6 THUẬT TOÁN KHAI PHÁ DỮ LIỆU')

add_math_paragraph(doc,
    'Trong chương này, nhóm tiến hành so sánh toàn diện **6 thuật toán khai phá dữ liệu** '
    'đã được học trong môn học. Các thuật toán được chia thành hai nhóm lớn:\n\n'
    '• **Nhóm 1 – Thuật toán Phân cụm (Clustering):** CURE, K-Means, K-Medoids — mục tiêu là '
    'gom các điểm dữ liệu tương đồng vào các cụm mà không cần nhãn trước.\n\n'
    '• **Nhóm 2 – Thuật toán Khai phá luật kết hợp (Association Rule Mining):** Apriori, FP-Growth, IT-Tree — '
    'mục tiêu là tìm kiếm các mẫu tần suất cao (frequent itemsets) và sinh ra các luật kết hợp '
    '$X \\Rightarrow Y$ trong cơ sở dữ liệu giao dịch.'
)

# ─────────────────────────────────────────────────────────────────
# 6.1  Phân nhóm và bài toán mỗi thuật toán giải quyết
# ─────────────────────────────────────────────────────────────────
add_heading_2('6.1 Phân nhóm và Bài toán từng thuật toán')

add_math_paragraph(doc,
    'Bảng 6.1 tóm tắt loại bài toán, đầu vào, đầu ra và ký hiệu '
    'tham số cơ bản của từng thuật toán.'
)

tbl1 = doc.add_table(rows=8, cols=5)
tbl1.alignment = WD_TABLE_ALIGNMENT.CENTER
set_table_borders(tbl1, '2E74B5')

headers1 = ['Thuật toán', 'Nhóm', 'Bài toán', 'Đầu vào chính', 'Đầu ra']
for ci, h in enumerate(headers1):
    cell = tbl1.rows[0].cells[ci]
    set_cell_background(cell, BLUE_HDR)
    p = cell.paragraphs[0]
    run = p.add_run(h)
    run.bold = True
    run.font.name = 'Times New Roman'
    run.font.size = Pt(11)
    run.font.color.rgb = RGBColor(255,255,255)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

rows1 = [
    ['CURE',       'Phân cụm', 'Gom cụm hình học phức tạp, kháng ngoại lai',
     'Tập điểm dữ liệu, k, c, α',
     'k cụm với đại diện co lại'],
    ['K-Means',    'Phân cụm', 'Gom cụm cầu cân đối, nhanh trên dữ liệu lớn',
     'Tập điểm, k',
     'k cụm với centroid μ_i'],
    ['K-Medoids',  'Phân cụm', 'Gom cụm bền vững với ngoại lai, dữ liệu phi số',
     'Tập điểm, k',
     'k cụm với medoid m_i'],
    ['Apriori',    'Khai phá luật kết hợp', 'Tìm tập phổ biến & luật X ⇒ Y theo BFS',
     'CSDL giao dịch, min_sup, min_conf',
     'Tập phổ biến, luật kết hợp'],
    ['FP-Growth',  'Khai phá luật kết hợp', 'Tìm tập phổ biến không sinh ứng viên',
     'CSDL giao dịch, min_sup',
     'Tập phổ biến, FP-tree'],
    ['IT-Tree',    'Khai phá luật kết hợp', 'Duyệt không gian itemset dạng cây trie tăng dần',
     'CSDL giao dịch, min_sup',
     'Tập phổ biến theo nhánh trie'],
]

for ri, row_data in enumerate(rows1):
    row = tbl1.rows[ri+1]
    bg = BLUE_LIGHT if ri % 2 == 0 else WHITE_ROW
    for ci, val in enumerate(row_data):
        cell = row.cells[ci]
        set_cell_background(cell, bg.replace('#',''))
        p = cell.paragraphs[0]
        if ci == 0:
            run = p.add_run(val)
            run.bold = True
            run.font.name = 'Times New Roman'; run.font.size = Pt(11)
        else:
            p.add_run(val).font.name = 'Times New Roman'
            p.runs[-1].font.size = Pt(11)

# ─────────────────────────────────────────────────────────────────
# 6.2  Nền tảng lý thuyết và công thức cốt lõi
# ─────────────────────────────────────────────────────────────────
add_heading_2('6.2 Nền tảng Lý thuyết và Công thức Cốt lõi')

# ───── 6.2.1 Nhóm Clustering ────────────────────────────────────
add_heading_3('6.2.1 Nhóm Phân cụm (Clustering)')

add_math_paragraph(doc,
    '**CURE (Clustering Using REpresentatives):** Mỗi cụm $C_i$ được đại diện bởi $c$ điểm '
    'đại diện $r_{i,1}, r_{i,2}, ..., r_{i,c}$ được co về phía centroid $\\mu_i$ theo hệ số $\\alpha \\in [0,1]$:'
)
add_display_equation(doc,
    "r'_{i,j} = r_{i,j} + \\alpha (\\mu_i - r_{i,j})",
    '(6.1)'
)
add_math_paragraph(doc,
    'Khoảng cách giữa hai cụm $C_a$ và $C_b$ là khoảng cách nhỏ nhất giữa các đại diện đã co:'
)
add_display_equation(doc,
    "d(C_a, C_b) = \\min_{r \\in C_a, s \\in C_b} d(r, s)",
    '(6.2)'
)

add_math_paragraph(doc,
    '**K-Means:** Thuật toán phân hoạch tối thiểu hóa tổng bình phương độ lệch '
    '(SSE — Sum of Squared Errors) nội cụm:'
)
add_display_equation(doc,
    "SSE = \\sum_{i=1}^{k} \\sum_{x \\in C_i} || x - \\mu_i ||^{2}",
    '(6.3)'
)
add_math_paragraph(doc, 'Trong đó $\\mu_i$ là centroid (trung bình cộng) của cụm $C_i$, được cập nhật sau mỗi vòng lặp theo:')
add_display_equation(doc,
    "\\mu_i = \\frac{1}{|C_i|} \\sum_{x \\in C_i} x",
    '(6.4)'
)

add_math_paragraph(doc,
    '**K-Medoids (PAM):** Thay vì centroid, mỗi cụm $C_i$ được đại diện bởi một điểm dữ liệu thực '
    '$m_i$ (medoid) sao cho tổng chi phí $J$ nhỏ nhất:'
)
add_display_equation(doc,
    "J = \\sum_{i=1}^{k} \\sum_{x \\in C_i} d(x, m_i)",
    '(6.5)'
)
add_math_paragraph(doc,
    'Thuật toán PAM thực hiện hoán đổi (swap) medoid $m_i$ với điểm $o \\in C_i$, '
    '$o \\neq m_i$ nếu chi phí giảm: $\\Delta J = J_{new} - J_{old} < 0$.'
)

# ───── 6.2.2 Nhóm Association Rule Mining ───────────────────────
add_heading_3('6.2.2 Nhóm Khai phá Luật kết hợp (Association Rule Mining)')

add_math_paragraph(doc,
    'Cho cơ sở dữ liệu giao dịch $D = \\{T_1, T_2, ..., T_n\\}$ gồm $n$ giao dịch, '
    'mỗi giao dịch $T_i \\subseteq I$ với $I$ là tập tất cả các mục (items). '
    'Các chỉ số đánh giá luật $X \\Rightarrow Y$ (với $X \\cap Y = \\emptyset$):\n\n'
    '• **Support** (độ phổ biến): tỷ lệ giao dịch chứa cả $X \\cup Y$:'
)
add_display_equation(doc,
    "support(X \\Rightarrow Y) = \\frac{|\\{T_i : X \\cup Y \\subseteq T_i\\}|}{n}",
    '(6.6)'
)
add_math_paragraph(doc, '• **Confidence** (độ tin cậy): xác suất có điều kiện:')
add_display_equation(doc,
    "confidence(X \\Rightarrow Y) = \\frac{support(X \\cup Y)}{support(X)}",
    '(6.7)'
)
add_math_paragraph(doc, '• **Lift** (độ nâng):')
add_display_equation(doc,
    "lift(X \\Rightarrow Y) = \\frac{confidence(X \\Rightarrow Y)}{support(Y)}",
    '(6.8)'
)

add_math_paragraph(doc,
    '**Apriori:** Áp dụng tính chất **Apriori** (tính chất phản đơn điệu): '
    'nếu một tập mục $X$ không phổ biến, mọi tập cha của $X$ đều không phổ biến. '
    'Thuật toán duyệt theo chiều rộng (BFS), lần lượt sinh các ứng viên '
    '$C_k$ từ các tập phổ biến $L_{k-1}$, rồi kiểm tra support của từng ứng viên bằng cách quét toàn bộ $D$.\n\n'
    '**Độ phức tạp:** $O(2^{|I|})$ trong trường hợp xấu nhất do số lượng ứng viên bùng nổ. '
    'Số lần quét CSDL bằng độ dài tập phổ biến dài nhất $k_{max}$.'
)

add_math_paragraph(doc,
    '**FP-Growth (Frequent Pattern Growth):** Nén CSDL thành cấu trúc **FP-tree** (cây tiền tố) '
    'chỉ với 2 lần quét toàn bộ $D$. Sau đó khai thác bằng đệ quy xây dựng '
    '**cơ sở mẫu điều kiện** (conditional pattern base) và **cây FP điều kiện** '
    'mà không sinh ứng viên tường minh. '
    'FP-Growth nhanh hơn Apriori từ **10 đến 100 lần** trên CSDL thưa.'
)

add_math_paragraph(doc,
    '**IT-Tree (Itemset Tree / Incremental Tidset Tree):** Lưu trữ không gian tập mục '
    'dưới dạng **cây trie** (prefix tree). Mỗi nút lưu một tập mục $X$ kèm '
    'danh sách định danh giao dịch (Tidset) $t(X)$. '
    'Từ đó $support(X) = |t(X)|$ và tập giao $t(X \\cup Y) = t(X) \\cap t(Y)$ '
    'được tính trực tiếp:\n'
)
add_display_equation(doc,
    "support(X \\cup Y) = |t(X) \\cap t(Y)|",
    '(6.9)'
)
add_math_paragraph(doc,
    'IT-Tree cho phép cập nhật tăng dần (incremental mining) khi CSDL thêm giao dịch mới '
    'mà không cần duyệt lại từ đầu, vì vậy phù hợp với dữ liệu luồng (data stream).'
)

# ─────────────────────────────────────────────────────────────────
# 6.3  So sánh toàn diện — Bảng tổng hợp
# ─────────────────────────────────────────────────────────────────
add_heading_2('6.3 Bảng So sánh Toàn diện 6 Thuật toán')

add_math_paragraph(doc, 'Bảng 6.2 so sánh các tiêu chí kỹ thuật quan trọng của cả 6 thuật toán.')

criteria = [
    'Tiêu chí',
    'CURE',
    'K-Means',
    'K-Medoids',
    'Apriori',
    'FP-Growth',
    'IT-Tree',
]

comp_rows = [
    ['Nhóm thuật toán',
     'Phân cụm phân cấp',
     'Phân cụm phân hoạch',
     'Phân cụm phân hoạch',
     'Khai phá luật kết hợp',
     'Khai phá luật kết hợp',
     'Khai phá luật kết hợp'],
    ['Chiến lược',
     'Đại diện co + Hierarchical',
     'EM-style, centroid',
     'PAM swap, medoid',
     'BFS, candidate gen.',
     'DFS, không sinh ứng viên',
     'Trie duyệt DFS + Tidset'],
    ['Tham số đầu vào',
     'k, c, α, s, p',
     'k',
     'k',
     'min_sup, min_conf',
     'min_sup',
     'min_sup'],
    ['Độ phức tạp thời gian',
     'O(n² log n)',
     'O(n·k·t·d)',
     'O(k(n−k)²)',
     'O(2^|I|) xấu nhất',
     'O(n·|FP-tree|)',
     'O(n·2^|I|) xấu nhất'],
    ['Độ phức tạp bộ nhớ',
     'O(n)',
     'O(n·k)',
     'O(n·k)',
     'O(2^|I|)',
     'O(n+|FP-tree|)',
     'O(n·|trie|)'],
    ['Quét CSDL',
     'Toàn bộ n điểm (lặp)',
     'Lặp đến hội tụ',
     'Lặp đến hội tụ',
     'k_max lần',
     '2 lần quét',
     '1–2 lần quét'],
    ['Kháng ngoại lai',
     '★★★★★ (tốt nhất)',
     '★★☆☆☆ (kém)',
     '★★★★☆ (tốt)',
     'N/A',
     'N/A',
     'N/A'],
    ['Hình dạng cụm',
     'Tùy ý, phi cầu',
     'Chỉ cầu cân đối',
     'Cầu, mạnh hơn K-Means',
     'N/A',
     'N/A',
     'N/A'],
    ['Cần xác định k trước',
     '✓ Có',
     '✓ Có',
     '✓ Có',
     '✗ Không',
     '✗ Không',
     '✗ Không'],
    ['Khả năng mở rộng\n(Scalability)',
     'Tốt (lấy mẫu s, p)',
     'Rất tốt',
     'Trung bình (PAM chậm)',
     'Kém (ứng viên bùng nổ)',
     'Tốt',
     'Tốt (incremental)'],
    ['Loại dữ liệu',
     'Số thực, không gian',
     'Số thực',
     'Bất kỳ (có hàm d)',
     'Giao dịch rời rạc',
     'Giao dịch rời rạc',
     'Giao dịch rời rạc'],
    ['Cập nhật tăng dần',
     '✗ Không',
     'Có thể (mini-batch)',
     '✗ Khó',
     '✗ Không (cần xử lý lại)',
     '✗ Không hoàn toàn',
     '✓ Có (thiết kế sẵn)'],
    ['Ứng dụng điển hình',
     'Phân tích không gian,\nảnh vệ tinh, gene',
     'Phân khúc khách hàng,\nnén ảnh, NLP',
     'Phân tích y tế,\ndữ liệu phức tạp',
     'Market basket,\nWeb usage mining',
     'Market basket quy mô\nlớn, text mining',
     'Data stream mining,\nDB động'],
]

ncols = 7
nrows = len(comp_rows) + 1
tbl2 = doc.add_table(rows=nrows, cols=ncols)
tbl2.alignment = WD_TABLE_ALIGNMENT.CENTER
set_table_borders(tbl2, '1F497D')

# header row
for ci, h in enumerate(criteria):
    cell = tbl2.rows[0].cells[ci]
    hdr_color = BLUE_HDR if ci == 0 else ('1E6B45' if ci >= 4 else BLUE_HDR)
    set_cell_background(cell, hdr_color)
    p = cell.paragraphs[0]
    run = p.add_run(h)
    run.bold = True; run.font.name = 'Times New Roman'
    run.font.size = Pt(10); run.font.color.rgb = RGBColor(255,255,255)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

for ri, row_data in enumerate(comp_rows):
    row = tbl2.rows[ri+1]
    for ci, val in enumerate(row_data):
        cell = row.cells[ci]
        if ci == 0:
            set_cell_background(cell, 'E8EEF7')
            p = cell.paragraphs[0]
            r = p.add_run(val); r.bold = True
            r.font.name = 'Times New Roman'; r.font.size = Pt(10)
        else:
            # Clustering cols = 1,2,3 → blue tint; ARM cols = 4,5,6 → green tint
            if ci <= 3:
                bg = BLUE_LIGHT if ri % 2 == 0 else WHITE_ROW
            else:
                bg = GREEN_LT if ri % 2 == 0 else WHITE_ROW
            set_cell_background(cell, bg.replace('#',''))
            p = cell.paragraphs[0]
            r = p.add_run(val)
            r.font.name = 'Times New Roman'; r.font.size = Pt(10)
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

# ─────────────────────────────────────────────────────────────────
# 6.4  So sánh theo từng tiêu chí quan trọng
# ─────────────────────────────────────────────────────────────────
add_heading_2('6.4 Phân tích Chi tiết Từng Tiêu chí')

add_heading_3('6.4.1 Độ phức tạp và Tốc độ')

add_math_paragraph(doc,
    'Trong **nhóm phân cụm**, K-Means là thuật toán nhanh nhất với độ phức tạp '
    '$O(n \\cdot k \\cdot t \\cdot d)$ (với $t$ là số vòng lặp, $d$ là chiều dữ liệu). '
    'K-Medoids chậm hơn đáng kể do bước swap của PAM có độ phức tạp $O(k(n-k)^{2})$. '
    'CURE có độ phức tạp $O(n^{2} \\log n)$ cho phiên bản cơ bản, nhưng được giảm xuống '
    'gần tuyến tính khi dùng lấy mẫu $s \\ll n$ điểm và phân hoạch $p$ tập con.\n\n'
    'Trong **nhóm khai phá luật kết hợp**, FP-Growth và IT-Tree vượt trội hơn Apriori do '
    'không sinh ứng viên tường minh. Apriori bị giới hạn bởi số lần quét CSDL '
    '$= k_{max}$ và số lượng ứng viên có thể lên tới $2^{|I|}$.'
)

add_heading_3('6.4.2 Chất lượng Kết quả')

add_math_paragraph(doc,
    'Tiêu chí đánh giá chất lượng khác nhau hoàn toàn giữa hai nhóm:\n\n'
    '**Nhóm phân cụm** — dùng **Silhouette Score** $s(x)$:\n'
)
add_display_equation(doc,
    "s(x) = \\frac{b(x) - a(x)}{\\max(a(x), b(x))}",
    '(6.10)'
)
add_math_paragraph(doc,
    'Trong đó $a(x)$ là khoảng cách trung bình đến các điểm cùng cụm (intra-cluster), '
    '$b(x)$ là khoảng cách trung bình đến cụm gần nhất khác (inter-cluster). '
    '$s(x) \\in [-1, 1]$, giá trị càng cao cụm càng tốt. '
    'Về Silhouette Score: **K-Means thường cho điểm cao nhất trên dữ liệu hình cầu**, nhưng '
    '**CURE cho điểm cao và đúng hơn trên dữ liệu phi cầu** (moons, rings, anisotropic).\n\n'
    '**Nhóm khai phá luật kết hợp** — đánh giá qua Support và Confidence (Công thức 6.6, 6.7, 6.8). '
    'Cả ba thuật toán Apriori, FP-Growth, IT-Tree **sinh ra cùng tập kết quả** khi dùng cùng '
    '$min\\_sup$. Sự khác biệt chỉ nằm ở **hiệu quả tính toán**, không phải chất lượng luật.'
)

add_heading_3('6.4.3 Khả năng xử lý Ngoại lai (Outliers)')

add_math_paragraph(doc,
    'Đây là tiêu chí phân biệt rõ nhất trong nhóm phân cụm:\n\n'
    '• **K-Means** — **Nhạy cảm nhất** với ngoại lai: centroid $\\mu_i$ bị kéo lệch '
    'do phép tính trung bình cộng cộng gộp tất cả điểm trong cụm, '
    'kể cả điểm ngoại lai cách xa rất nhiều.\n\n'
    '• **K-Medoids** — **Bền vững hơn K-Means**: medoid $m_i$ là điểm dữ liệu thực '
    'nên không bị kéo lệch về phía ngoại lai. Tuy nhiên, khi $k$ nhỏ, '
    'ngoại lai vẫn có thể trở thành medoid.\n\n'
    '• **CURE** — **Bền vững nhất**: CURE phát hiện và loại bỏ ngoại lai trong giai đoạn '
    'tiền xử lý (cụm có ít hơn $n/k \\cdot \\varepsilon$ điểm bị đánh dấu là ngoại lai). '
    'Cơ chế co $\\alpha$ cũng giúp giảm ảnh hưởng của các điểm ngoại biên cụm.'
)

add_heading_3('6.4.4 Khả năng mở rộng và Dữ liệu lớn')

add_math_paragraph(doc,
    'Khi $n$ tăng lên hàng triệu bản ghi:\n\n'
    '• **K-Means** — Mở rộng tốt nhất trong nhóm phân cụm. Mini-Batch K-Means '
    'xử lý từng batch $B \\subset D$ với $|B| \\ll n$ cho phép chạy trên dữ liệu lớn.\n\n'
    '• **CURE** — Dùng lấy mẫu ngẫu nhiên $s$ điểm từ $D$, phân hoạch thành $p$ tập con, '
    'phân cụm độc lập rồi gộp lại. Độ chính xác phụ thuộc vào $s$ và $p$.\n\n'
    '• **K-Medoids (PAM)** — Kém mở rộng nhất trong nhóm phân cụm do swap $O(k(n-k)^{2})$. '
    'Biến thể CLARA và CLARANS cải thiện bằng lấy mẫu.\n\n'
    '• **FP-Growth** — Mở rộng tốt trong nhóm ARM; '
    'kích thước FP-tree thường nhỏ hơn CSDL gốc nhiều lần nhờ nén tiền tố chung.\n\n'
    '• **Apriori** — Kém mở rộng do số ứng viên bùng nổ theo hàm mũ $2^{|I|}$.\n\n'
    '• **IT-Tree** — Phù hợp dữ liệu luồng và CSDL cập nhật liên tục nhờ thiết kế incremental.'
)

# ─────────────────────────────────────────────────────────────────
# 6.5  So sánh các bước thuật toán (Pseudocode tóm tắt)
# ─────────────────────────────────────────────────────────────────
add_heading_2('6.5 Lưu đồ Bước Thực thi Tổng quan')

add_math_paragraph(doc,
    'Bảng 6.3 tóm tắt các bước chính (pseudocode rút gọn) của từng thuật toán để '
    'tiện so sánh luồng xử lý.'
)

pseudo_rows = [
    ['CURE',
     '1. Lấy mẫu s điểm\n2. Phân hoạch thành p tập\n3. Phân cụm trực tiếp từng tập\n4. Loại ngoại lai\n5. Gộp cụm → chọn c đại diện → co α\n6. Gán điểm còn lại → k cụm'],
    ['K-Means',
     '1. Chọn k centroid ngẫu nhiên\n2. Gán mỗi x ∈ D vào cụm gần nhất\n3. Cập nhật μ_i = mean(C_i)\n4. Lặp 2–3 đến hội tụ (ΔSSE < ε)'],
    ['K-Medoids',
     '1. Chọn k medoid ngẫu nhiên m_i\n2. Gán x vào cụm có m_i gần nhất\n3. Với mỗi m_i, thử swap với o ∈ C_i\n4. Giữ swap nếu ΔJ < 0\n5. Lặp 2–4 đến không còn swap cải thiện'],
    ['Apriori',
     '1. Quét D → L_1 (frequent 1-itemsets)\n2. C_k = generate_candidates(L_{k-1})\n3. Quét D → tính support(c) ∀ c ∈ C_k\n4. L_k = {c ∈ C_k : support(c) ≥ min_sup}\n5. Lặp 2–4 đến C_k rỗng\n6. Sinh luật từ ∪L_k'],
    ['FP-Growth',
     '1. Quét D lần 1 → F-list (frequent items)\n2. Quét D lần 2 → xây FP-tree\n3. Với item p: trích conditional pattern base\n4. Xây conditional FP-tree → khai thác đệ quy\n5. Sinh tập phổ biến, không sinh ứng viên'],
    ['IT-Tree',
     '1. Xây dựng trie rỗng\n2. Duyệt DFS từng nút X của trie\n3. Tính t(X) = ∩ t(item) cho mỗi item ∈ X\n4. support(X) = |t(X)| ≥ min_sup → lưu X\n5. Mở rộng nút X → X∪{item} (prefix sharing)\n6. Cập nhật tăng dần khi D thêm giao dịch'],
]

tbl3 = doc.add_table(rows=len(pseudo_rows)+1, cols=2)
tbl3.alignment = WD_TABLE_ALIGNMENT.CENTER
tbl3.columns[0].width = Cm(3.5)
tbl3.columns[1].width = Cm(12.5)
set_table_borders(tbl3, '1F497D')

for ci, h in enumerate(['Thuật toán', 'Các bước chính']):
    cell = tbl3.rows[0].cells[ci]
    set_cell_background(cell, BLUE_HDR)
    p = cell.paragraphs[0]
    r = p.add_run(h); r.bold = True
    r.font.name = 'Times New Roman'; r.font.size = Pt(11)
    r.font.color.rgb = RGBColor(255,255,255)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

for ri, (algo, steps) in enumerate(pseudo_rows):
    row = tbl3.rows[ri+1]
    bg  = BLUE_LIGHT if ri < 3 else GREEN_LT
    # algo name
    c0 = row.cells[0]
    set_cell_background(c0, bg.replace('#',''))
    r = c0.paragraphs[0].add_run(algo)
    r.bold = True; r.font.name = 'Times New Roman'; r.font.size = Pt(11)
    # steps
    c1 = row.cells[1]
    set_cell_background(c1, WHITE_ROW)
    r1 = c1.paragraphs[0].add_run(steps)
    r1.font.name = 'Courier New'; r1.font.size = Pt(10)

# ─────────────────────────────────────────────────────────────────
# 6.6  Khi nào nên dùng thuật toán nào?
# ─────────────────────────────────────────────────────────────────
add_heading_2('6.6 Hướng dẫn Lựa chọn Thuật toán Phù hợp')

add_math_paragraph(doc,
    'Việc lựa chọn thuật toán phụ thuộc vào bản chất bài toán và đặc điểm dữ liệu:\n\n'
    '**→ Dùng CURE khi:**\n'
    '   • Dữ liệu có cụm hình dạng tùy ý (không phải hình cầu cân đối).\n'
    '   • Dữ liệu chứa nhiều ngoại lai cần phát hiện và loại bỏ.\n'
    '   • Kích thước dữ liệu lớn nhưng cần kết quả chính xác cao ($k, c, \\alpha$ được tinh chỉnh).\n\n'
    '**→ Dùng K-Means khi:**\n'
    '   • Dữ liệu có cụm hình cầu, cân đối, kích thước tương đương.\n'
    '   • Ưu tiên tốc độ và đơn giản; $n$ rất lớn (hàng triệu điểm).\n\n'
    '**→ Dùng K-Medoids khi:**\n'
    '   • Dữ liệu không phải số thực, hoặc hàm khoảng cách không phải Euclidean.\n'
    '   • Cần bền vững hơn K-Means với ngoại lai nhưng không cần độ phức tạp của CURE.\n\n'
    '**→ Dùng Apriori khi:**\n'
    '   • Tập item $|I|$ nhỏ hoặc vừa, $min\\_sup$ cao → số ứng viên không bùng nổ.\n'
    '   • Cần cài đặt đơn giản, dễ debug, dễ giải thích từng bước.\n\n'
    '**→ Dùng FP-Growth khi:**\n'
    '   • CSDL giao dịch lớn, $|I|$ nhiều, $min\\_sup$ thấp → Apriori không khả thi.\n'
    '   • Đây là lựa chọn mặc định cho hầu hết bài toán ARM thực tế.\n\n'
    '**→ Dùng IT-Tree khi:**\n'
    '   • CSDL thay đổi liên tục (thêm/xóa giao dịch) — cần cập nhật tăng dần.\n'
    '   • Phù hợp với hệ thống thời gian thực (data stream, e-commerce động).'
)

# ─────────────────────────────────────────────────────────────────
# 6.7  Tổng kết chương
# ─────────────────────────────────────────────────────────────────
add_heading_2('6.7 Tổng kết Chương 6')

add_math_paragraph(doc,
    'Qua phân tích so sánh 6 thuật toán khai phá dữ liệu, ta rút ra nhận xét tổng quát:\n\n'
    '1. **Hai nhóm bổ sung cho nhau:** Phân cụm (CURE, K-Means, K-Medoids) tìm cấu trúc '
    'tiềm ẩn trong dữ liệu không nhãn, trong khi Khai phá luật kết hợp (Apriori, FP-Growth, IT-Tree) '
    'tìm mối liên hệ giữa các mục trong giao dịch. Không có thuật toán nào "tốt hơn tuyệt đối" '
    '— mỗi thuật toán được thiết kế cho một bài toán riêng biệt.\n\n'
    '2. **Trong nhóm phân cụm:** CURE là thuật toán mạnh nhất về khả năng xử lý hình dạng tùy ý '
    'và kháng ngoại lai. K-Means nhanh nhất, K-Medoids bền vững nhất với metric tùy ý.\n\n'
    '3. **Trong nhóm ARM:** FP-Growth thực tế hiệu quả nhất. Apriori đơn giản nhất để học. '
    'IT-Tree độc đáo nhất cho dữ liệu động.\n\n'
    '4. **Độ phức tạp là yếu tố then chốt:** Khi $n$ và $|I|$ tăng, sự khác biệt về độ phức tạp '
    'từ lý thuyết $O(\\cdot)$ trở nên rất rõ ràng trong thực tế. '
    'Việc lựa chọn đúng thuật toán có thể rút ngắn thời gian xử lý từ hàng giờ xuống vài giây.'
)

# ─────────────────────────────────────────────────────────────────
# Lưu file
# ─────────────────────────────────────────────────────────────────
doc.save(OUTPUT_FILE)
print(f"✅ Đã bổ sung Chương 6 thành công vào: {OUTPUT_FILE}")
