import os
import shutil
from PIL import Image, ImageDraw, ImageFont
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
FONT_DIR = r"C:\Windows\Fonts"

FONT_REGULAR = os.path.join(FONT_DIR, "segoeui.ttf")
FONT_BOLD = os.path.join(FONT_DIR, "segoeuib.ttf")
FONT_ITALIC = os.path.join(FONT_DIR, "segoeuii.ttf")
FONT_MONO = os.path.join(FONT_DIR, "consola.ttf")


def load_font(path, size):
    if os.path.exists(path):
        return ImageFont.truetype(path, size)
    return ImageFont.load_default()


def draw_rounded_box(draw, xy, radius, fill, outline, width=3):
    x0, y0, x1, y1 = xy
    # Shadow
    draw.rounded_rectangle(
        [x0 + 6, y0 + 6, x1 + 6, y1 + 6],
        radius=radius,
        fill=(226, 232, 240)
    )
    draw.rounded_rectangle(
        [x0, y0, x1, y1],
        radius=radius,
        fill=fill,
        outline=outline,
        width=width
    )


def draw_header_box(draw, xy, radius, header_h, header_fill, body_fill, outline, title, subtitle_lines, title_font, body_font, mono_font=None):
    x0, y0, x1, y1 = xy
    draw_rounded_box(draw, xy, radius, body_fill, outline, width=3)
    # Header fill (top rounded, bottom flat)
    draw.rounded_rectangle([x0, y0, x1, y0 + header_h], radius=radius, fill=header_fill, outline=outline, width=3)
    draw.rectangle([x0 + 2, y0 + header_h - radius, x1 - 2, y0 + header_h], fill=header_fill)
    draw.line([x0, y0 + header_h, x1, y0 + header_h], fill=outline, width=2)

    # Title centered in header
    bbox = draw.textbbox((0, 0), title, font=title_font)
    tw = bbox[2] - bbox[0]
    th = bbox[3] - bbox[1]
    draw.text(((x0 + x1 - tw) / 2, y0 + (header_h - th) / 2 - 4), title, font=title_font, fill=(255, 255, 255))

    # Body lines
    curr_y = y0 + header_h + 18
    for line_text, color, is_mono in subtitle_lines:
        f = mono_font if (is_mono and mono_font) else body_font
        draw.text((x0 + 20, curr_y), line_text, font=f, fill=color)
        curr_y += 32


def draw_arrow(draw, points, color, width=4, arrow_size=14):
    for i in range(len(points) - 1):
        draw.line([points[i], points[i + 1]], fill=color, width=width)
    x1, y1 = points[-2]
    x2, y2 = points[-1]
    dx = x2 - x1
    dy = y2 - y1
    length = (dx ** 2 + dy ** 2) ** 0.5
    if length == 0:
        return
    ux, uy = dx / length, dy / length
    px, py = -uy, ux
    p1 = (x2, y2)
    p2 = (x2 - ux * arrow_size + px * (arrow_size * 0.55), y2 - uy * arrow_size + py * (arrow_size * 0.55))
    p3 = (x2 - ux * arrow_size - px * (arrow_size * 0.55), y2 - uy * arrow_size - py * (arrow_size * 0.55))
    draw.polygon([p1, p2, p3], fill=color)


def draw_label_pill(draw, center_x, center_y, text, bg_color, border_color, text_color, font):
    bbox = draw.textbbox((0, 0), text, font=font)
    tw = bbox[2] - bbox[0]
    th = bbox[3] - bbox[1]
    pad_x, pad_y = 14, 8
    x0 = center_x - tw / 2 - pad_x
    y0 = center_y - th / 2 - pad_y
    x1 = center_x + tw / 2 + pad_x
    y1 = center_y + th / 2 + pad_y
    draw.rounded_rectangle([x0, y0, x1, y1], radius=10, fill=bg_color, outline=border_color, width=2)
    draw.text((center_x - tw / 2, center_y - th / 2 - 3), text, font=font, fill=text_color)


def generate_component_tree_diagram(out_path):
    W, H = 1900, 1080
    img = Image.new("RGB", (W, H), (248, 250, 252))
    draw = ImageDraw.Draw(img)

    f_title = load_font(FONT_BOLD, 34)
    f_sub = load_font(FONT_REGULAR, 22)
    f_box_title = load_font(FONT_BOLD, 26)
    f_body = load_font(FONT_REGULAR, 21)
    f_mono = load_font(FONT_MONO, 19)
    f_pill = load_font(FONT_BOLD, 18)

    # Title
    t_str = "SƠ ĐỒ CÂY COMPONENT (COMPONENT TREE) - ỨNG DỤNG REACT CRUD"
    tb = draw.textbbox((0, 0), t_str, font=f_title)
    draw.text(((W - (tb[2] - tb[0])) / 2, 34), t_str, font=f_title, fill=(15, 23, 42))
    sub_str = "Kiến trúc phân tầng Component & Quản lý State tập trung tại Root Component (App)"
    sb = draw.textbbox((0, 0), sub_str, font=f_sub)
    draw.text(((W - (sb[2] - sb[0])) / 2, 82), sub_str, font=f_sub, fill=(71, 85, 105))

    # Root DOM node
    draw_rounded_box(draw, (740, 135, 1160, 200), 12, (241, 245, 249), (100, 116, 139), width=2)
    dom_str = 'ReactDOM.createRoot(#root)'
    db = draw.textbbox((0, 0), dom_str, font=f_mono)
    draw.text(((W - (db[2] - db[0])) / 2, 155), dom_str, font=f_mono, fill=(51, 65, 85))

    # Arrow Root -> App
    draw_arrow(draw, [(950, 200), (950, 245)], (100, 116, 139), width=3)

    # App node
    draw_header_box(
        draw,
        xy=(580, 245, 1320, 465),
        radius=16,
        header_h=56,
        header_fill=(30, 58, 95),
        body_fill=(239, 246, 255),
        outline=(37, 99, 235),
        title="<App />  (Root Component)",
        subtitle_lines=[
            ("• State dùng chung:", (30, 58, 95), False),
            ("    const [kw, setKeyword] = useState(\"\")", (29, 78, 216), True),
            ("    const [newUser, setNewUser] = useState(null)", (29, 78, 216), True),
            ("• Vai trò: Lưu trữ state tập trung & điều phối Props", (51, 65, 85), False),
        ],
        title_font=f_box_title,
        body_font=f_body,
        mono_font=f_mono
    )

    # Tree connectors from App (x=950, y=465) down to horizontal bus at y=535
    draw.line([(950, 465), (950, 535)], fill=(51, 65, 85), width=4)
    draw.line([(310, 535), (1570, 535)], fill=(51, 65, 85), width=4)
    draw_arrow(draw, [(310, 535), (310, 620)], (15, 118, 110), width=4)
    draw_arrow(draw, [(950, 535), (950, 620)], (180, 83, 9), width=4)
    draw_arrow(draw, [(1570, 535), (1570, 620)], (67, 56, 202), width=4)

    # Prop pills on branches
    draw_label_pill(draw, 310, 575, "prop: onChangeValue={setKeyword}", (240, 253, 250), (20, 184, 166), (15, 118, 110), f_pill)
    draw_label_pill(draw, 950, 575, "prop: onAdd={setNewUser}", (255, 251, 235), (245, 158, 11), (180, 83, 9), f_pill)
    draw_label_pill(draw, 1570, 575, "props: keyword, user, onAdded", (238, 242, 255), (99, 102, 241), (67, 56, 202), f_pill)

    # Child 1: SearchForm
    draw_header_box(
        draw,
        xy=(60, 620, 560, 930),
        radius=16,
        header_h=54,
        header_fill=(15, 118, 110),
        body_fill=(240, 253, 250),
        outline=(20, 184, 166),
        title="<SearchForm />",
        subtitle_lines=[
            ("• Props nhận:", (15, 118, 110), False),
            ("    { onChangeValue }", (17, 94, 89), True),
            ("• State nội bộ: Không (Stateless)", (51, 65, 85), False),
            ("• Chức năng:", (15, 118, 110), False),
            ("    Render ô <input> tìm kiếm", (51, 65, 85), False),
            ("    Gửi giá trị gõ lên App qua callback", (51, 65, 85), False),
        ],
        title_font=f_box_title,
        body_font=f_body,
        mono_font=f_mono
    )

    # Child 2: AddUser
    draw_header_box(
        draw,
        xy=(620, 620, 1280, 930),
        radius=16,
        header_h=54,
        header_fill=(180, 83, 9),
        body_fill=(255, 251, 235),
        outline=(245, 158, 11),
        title="<AddUser />",
        subtitle_lines=[
            ("• Props nhận: { onAdd }", (180, 83, 9), False),
            ("• State nội bộ:", (180, 83, 9), False),
            ("    adding (boolean), user (object)", (146, 64, 14), True),
            ("• Chức năng:", (180, 83, 9), False),
            ("    Nút '+ Thêm người dùng' mở Modal", (51, 65, 85), False),
            ("    Xử lý nested address (...spread) & submit", (51, 65, 85), False),
        ],
        title_font=f_box_title,
        body_font=f_body,
        mono_font=f_mono
    )

    # Child 3: ResultTable
    draw_header_box(
        draw,
        xy=(1330, 620, 1840, 930),
        radius=16,
        header_h=54,
        header_fill=(67, 56, 202),
        body_fill=(238, 242, 255),
        outline=(99, 102, 241),
        title="<ResultTable />",
        subtitle_lines=[
            ("• Props: { keyword, user, onAdded }", (67, 56, 202), False),
            ("• State nội bộ:", (67, 56, 202), False),
            ("    users[], loading, editing", (55, 48, 163), True),
            ("• Chức năng:", (67, 56, 202), False),
            ("    useEffect() fetch API & thêm user", (51, 65, 85), False),
            ("    Lọc, Sửa (Deep Copy Modal), Xóa", (51, 65, 85), False),
        ],
        title_font=f_box_title,
        body_font=f_body,
        mono_font=f_mono
    )

    # Footer note
    draw_rounded_box(draw, (60, 968, 1840, 1035), 10, (241, 245, 249), (203, 213, 225), width=2)
    note = "Ghi chú kiến trúc: Cây component có độ sâu 2 cấp (App -> 3 Sibling Components). State chia sẻ được đặt tại App; các state cục bộ UI đặt tại từng con."
    nb = draw.textbbox((0, 0), note, font=f_body)
    draw.text(((W - (nb[2] - nb[0])) / 2, 987), note, font=f_body, fill=(51, 65, 85))

    img.save(out_path, dpi=(300, 300))


def generate_state_lifting_diagram(out_path):
    W, H = 1960, 1140
    img = Image.new("RGB", (W, H), (248, 250, 252))
    draw = ImageDraw.Draw(img)

    f_title = load_font(FONT_BOLD, 34)
    f_sub = load_font(FONT_REGULAR, 22)
    f_box_title = load_font(FONT_BOLD, 25)
    f_body = load_font(FONT_REGULAR, 20)
    f_mono = load_font(FONT_MONO, 19)
    f_pill = load_font(FONT_BOLD, 18)

    # Title
    t_str = "SƠ ĐỒ LUỒNG DỮ LIỆU (STATE LIFTING & ONE-WAY DATA FLOW)"
    tb = draw.textbbox((0, 0), t_str, font=f_title)
    draw.text(((W - (tb[2] - tb[0])) / 2, 30), t_str, font=f_title, fill=(15, 23, 42))
    sub_str = "Dữ liệu truyền lên Cha qua Callback Props, sau đó truyền xuống Con nhận kết quả qua Data Props"
    sb = draw.textbbox((0, 0), sub_str, font=f_sub)
    draw.text(((W - (sb[2] - sb[0])) / 2, 78), sub_str, font=f_sub, fill=(71, 85, 105))

    # Top Center: App Component
    draw_header_box(
        draw,
        xy=(490, 145, 1450, 385),
        radius=18,
        header_h=56,
        header_fill=(30, 58, 95),
        body_fill=(239, 246, 255),
        outline=(37, 99, 235),
        title="<App />  (Single Source of Truth - Nơi giữ State tập trung)",
        subtitle_lines=[
            ("1. State Tìm kiếm:  const [kw, setKeyword] = React.useState(\"\")", (15, 118, 110), True),
            ("2. State Thêm mới:  const [newUser, setNewUser] = React.useState(null)", (37, 99, 235), True),
            ("-> Nhận callback từ con, cập nhật state và truyền props mới xuống ResultTable", (51, 65, 85), False),
        ],
        title_font=f_box_title,
        body_font=f_body,
        mono_font=f_mono
    )

    # Bottom Left: SearchForm
    draw_header_box(
        draw,
        xy=(60, 560, 550, 815),
        radius=16,
        header_h=52,
        header_fill=(5, 150, 105),
        body_fill=(236, 253, 245),
        outline=(16, 185, 129),
        title="<SearchForm />  (Nguồn từ khóa)",
        subtitle_lines=[
            ("• Người dùng nhập vào ô <input>", (30, 41, 59), False),
            ("• Sự kiện onChange kích hoạt:", (30, 41, 59), False),
            ("  onChangeValue(e.target.value)", (4, 120, 87), True),
            ("• Gọi trực tiếp setKeyword ở App", (71, 85, 105), False),
        ],
        title_font=f_box_title,
        body_font=f_body,
        mono_font=f_mono
    )

    # Bottom Center: AddUser
    draw_header_box(
        draw,
        xy=(600, 680, 1105, 945),
        radius=16,
        header_h=52,
        header_fill=(217, 119, 6),
        body_fill=(255, 251, 235),
        outline=(245, 158, 11),
        title="<AddUser />  (Nguồn User mới)",
        subtitle_lines=[
            ("• Nhập form & lưu state cục bộ user", (30, 41, 59), False),
            ("• Nhấn 'Thêm' -> gọi callback:", (30, 41, 59), False),
            ("  onAdd(user)", (180, 83, 9), True),
            ("• Gọi setNewUser(user) ở App", (71, 85, 105), False),
        ],
        title_font=f_box_title,
        body_font=f_body,
        mono_font=f_mono
    )

    # Bottom Right: ResultTable
    draw_header_box(
        draw,
        xy=(1360, 540, 1910, 915),
        radius=16,
        header_h=52,
        header_fill=(67, 56, 202),
        body_fill=(238, 242, 255),
        outline=(99, 102, 241),
        title="<ResultTable />  (Tiêu thụ dữ liệu)",
        subtitle_lines=[
            ("• Nhận keyword -> lọc danh sách:", (30, 41, 59), False),
            ("  users.filter(u => ...keyword)", (55, 48, 163), True),
            ("• Nhận user -> useEffect([user]):", (30, 41, 59), False),
            ("  Thêm user vào state users[]", (55, 48, 163), True),
            ("• Gọi onAdded() -> reset newUser = null", (124, 58, 237), True),
            ("• Quản lý Edit (Deep Copy) & Delete", (71, 85, 105), False),
        ],
        title_font=f_box_title,
        body_font=f_body,
        mono_font=f_mono
    )

    # Arrow 1: SearchForm -> App (Lifting kw up)
    draw_arrow(draw, [(305, 560), (305, 260), (490, 260)], (5, 150, 105), width=5)
    draw_label_pill(draw, 305, 410, "(1) Callback: onChangeValue(kw)", (236, 253, 245), (5, 150, 105), (4, 120, 87), f_pill)

    # Arrow 2: AddUser -> App (Lifting newUser up)
    draw_arrow(draw, [(850, 680), (850, 385)], (217, 119, 6), width=5)
    draw_label_pill(draw, 850, 515, "(3) Callback: onAdd(user)", (255, 251, 235), (217, 119, 6), (180, 83, 9), f_pill)

    # Arrow 3: App -> ResultTable (Flow 1: keyword={kw})
    draw_arrow(draw, [(1450, 225), (1710, 225), (1710, 540)], (5, 150, 105), width=5)
    draw_label_pill(draw, 1710, 315, "(2) Prop: keyword={kw}", (236, 253, 245), (5, 150, 105), (4, 120, 87), f_pill)

    # Arrow 4: App -> ResultTable (Flow 2: user={newUser})
    draw_arrow(draw, [(1450, 320), (1525, 320), (1525, 540)], (37, 99, 235), width=5)
    draw_label_pill(draw, 1525, 430, "(4) Prop: user={newUser}", (239, 246, 255), (37, 99, 235), (29, 78, 216), f_pill)

    # Arrow 5: ResultTable -> App (Reset newUser via onAdded())
    draw_arrow(draw, [(1360, 605), (1205, 605), (1205, 385)], (124, 58, 237), width=4)
    draw_label_pill(draw, 1205, 505, "(5) onAdded() -> reset null", (245, 243, 255), (124, 58, 237), (109, 40, 217), f_pill)

    # Forbidden direct sibling arrow (dashed red X) between SearchForm/AddUser and ResultTable?
    # Better: show a clear Legend & Sibling Isolation Note at bottom
    draw_rounded_box(draw, (70, 990, 1900, 1100), 12, (254, 242, 242), (239, 68, 68), width=2)
    w1 = "TẠI SAO KHÔNG ĐẶT STATE Ở CON?  SearchForm, AddUser và ResultTable là các Sibling Components (anh em cùng cấp)."
    w2 = "React tuân theo One-Way Data Flow (chỉ truyền Cha -> Con): các con KHÔNG thể truyền dữ liệu trực tiếp sang cho nhau -> Phải nâng state (State Lifting) lên App."
    wb1 = draw.textbbox((0, 0), w1, font=load_font(FONT_BOLD, 21))
    wb2 = draw.textbbox((0, 0), w2, font=f_body)
    draw.text(((W - (wb1[2] - wb1[0])) / 2, 1008), w1, font=load_font(FONT_BOLD, 21), fill=(185, 28, 28))
    draw.text(((W - (wb2[2] - wb2[0])) / 2, 1048), w2, font=f_body, fill=(127, 29, 29))

    img.save(out_path, dpi=(300, 300))


# ================== WORD HELPER FUNCTIONS ==================
def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_hex)
    tcPr.append(shd)


def set_cell_margins(cell, top=100, bottom=100, left=140, right=140):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)


def set_table_borders(table, border_color="CBD5E1"):
    tblPr = table._tbl.tblPr
    borders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '4')
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), border_color)
        borders.append(border)
    tblPr.append(borders)


def add_heading_1(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run(text)
    r.font.name = "Arial"
    r.font.size = Pt(14)
    r.font.bold = True
    r.font.color.rgb = RGBColor(30, 58, 95)
    return p


def add_heading_2(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(text)
    r.font.name = "Arial"
    r.font.size = Pt(11.5)
    r.font.bold = True
    r.font.color.rgb = RGBColor(37, 99, 235)
    return p


def add_body_p(doc, text, bold_prefix=None, space_after=5):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.25
    if bold_prefix:
        rb = p.add_run(bold_prefix)
        rb.font.name = "Arial"
        rb.font.size = Pt(10.5)
        rb.font.bold = True
    r = p.add_run(text)
    r.font.name = "Arial"
    r.font.size = Pt(10.5)
    return p


def add_code_block(doc, code_text):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.rows[0].cells[0]
    cell.width = Inches(6.5)
    set_cell_background(cell, "F1F5F9")
    set_cell_margins(cell, top=110, bottom=110, left=160, right=160)
    set_table_borders(tbl, "CBD5E1")
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.15
    r = p.add_run(code_text)
    r.font.name = "Consolas"
    r.font.size = Pt(9)
    r.font.color.rgb = RGBColor(30, 41, 59)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def add_styled_table(doc, headers, rows, col_widths, center_cols=None):
    if center_cols is None:
        center_cols = []
    tbl = doc.add_table(rows=1, cols=len(headers))
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    set_table_borders(tbl, "CBD5E1")

    hdr_cells = tbl.rows[0].cells
    for i, h in enumerate(headers):
        hdr_cells[i].text = h
        set_cell_background(hdr_cells[i], "1E3A5F")
        set_cell_margins(hdr_cells[i], top=110, bottom=110, left=110, right=110)
        hdr_cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(0)
        for run in p.runs:
            run.font.name = "Arial"
            run.font.size = Pt(9.5)
            run.font.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)

    for r_idx, row_data in enumerate(rows):
        cells = tbl.add_row().cells
        bg = "F8FAFC" if r_idx % 2 == 1 else "FFFFFF"
        for c_idx, val in enumerate(row_data):
            cells[c_idx].text = val
            set_cell_background(cells[c_idx], bg)
            set_cell_margins(cells[c_idx], top=85, bottom=85, left=105, right=105)
            cells[c_idx].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cells[c_idx].paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.15
            if c_idx in center_cols:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                run.font.name = "Consolas" if (val.startswith("on") or val in ["keyword", "user", "kw", "newUser", "setKeyword", "setNewUser", "() => setNewUser(null)"]) else "Arial"
                run.font.size = Pt(9.2)
                if c_idx == 1 and len(headers) >= 5:
                    run.font.bold = True
                    run.font.color.rgb = RGBColor(29, 78, 216)

    for row in tbl.rows:
        for i, w in enumerate(col_widths):
            row.cells[i].width = w
    return tbl


def build_report():
    tree_img_path = os.path.join(CURRENT_DIR, "component_tree.png")
    flow_img_path = os.path.join(CURRENT_DIR, "state_lifting_flow.png")
    generate_component_tree_diagram(tree_img_path)
    generate_state_lifting_diagram(flow_img_path)

    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = Inches(0.85)
        section.bottom_margin = Inches(0.85)
        section.left_margin = Inches(0.9)
        section.right_margin = Inches(0.9)

    # Header / Title Banner
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_p.paragraph_format.space_after = Pt(4)
    rt = title_p.add_run("BÁO CÁO KẾT QUẢ BÀI TẬP: REACT CƠ BẢN (CRUD)")
    rt.font.name = "Arial"
    rt.font.size = Pt(18)
    rt.font.bold = True
    rt.font.color.rgb = RGBColor(30, 58, 95)

    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub_p.paragraph_format.space_after = Pt(6)
    rs = sub_p.add_run("Ứng dụng Quản lý Người dùng (Create - Read - Update - Delete) với React 18, Hooks & State Lifting")
    rs.font.name = "Arial"
    rs.font.size = Pt(11)
    rs.font.italic = True
    rs.font.color.rgb = RGBColor(71, 85, 105)

    meta_p = doc.add_paragraph()
    meta_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    meta_p.paragraph_format.space_after = Pt(14)
    rm = meta_p.add_run("Học phần: IT4409 - Công nghệ Web và Dịch vụ trực tuyến  |  Sinh viên thực hiện: Vũ Hòa Bình (MSSV: 202416137)")
    rm.font.name = "Arial"
    rm.font.size = Pt(10)
    rm.font.bold = True
    rm.font.color.rgb = RGBColor(37, 99, 235)

    # ==================== PHẦN 1: SƠ ĐỒ COMPONENT ====================
    add_heading_1(doc, "1. SƠ ĐỒ CÂY COMPONENT (COMPONENT TREE)")
    add_body_p(
        doc,
        "Ứng dụng Quản lý người dùng được xây dựng theo kiến trúc Functional Component của React 18, gồm 1 component gốc (App) và 3 component con ngang hàng (SearchForm, AddUser, ResultTable) được render vào thẻ <div id=\"root\"> thông qua ReactDOM.createRoot()."
    )

    # Insert Component Tree Image
    pic_p1 = doc.add_paragraph()
    pic_p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pic_p1.paragraph_format.space_after = Pt(4)
    pic_p1.add_run().add_picture(tree_img_path, width=Inches(6.5))

    cap1 = doc.add_paragraph()
    cap1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap1.paragraph_format.space_after = Pt(8)
    rc1 = cap1.add_run("Hình 1.1: Sơ đồ cây Component (Component Tree) của ứng dụng React CRUD.")
    rc1.font.name = "Arial"
    rc1.font.size = Pt(9.5)
    rc1.font.italic = True
    rc1.font.color.rgb = RGBColor(100, 116, 139)

    add_code_block(doc, """index.html (<div id="root"></div>)
└── <App />  [State dùng chung: kw, newUser]
    ├── <SearchForm onChangeValue={setKeyword} />        [Stateless - nhận input tìm kiếm]
    ├── <AddUser onAdd={setNewUser} />                   [State nội bộ: adding, user]
    └── <ResultTable keyword={kw} user={newUser} onAdded={() => setNewUser(null)} />
                                                         [State nội bộ: users, loading, error, editing]""")

    comp_headers = ["Component", "Loại / Cấp", "State nội bộ (Local State)", "Vai trò chính trong ứng dụng"]
    comp_rows = [
        ("App", "Root (Cha)", "kw (string), newUser (object | null)", "Lưu trữ các state chia sẻ giữa các component con; truyền callback và dữ liệu xuống thông qua Props."),
        ("SearchForm", "Child (Con)", "Không có (Stateless Component)", "Hiển thị ô <input> tìm kiếm; kích hoạt onChangeValue(e.target.value) mỗi khi người dùng gõ phím."),
        ("AddUser", "Child (Con)", "adding (boolean), user (object lồng address)", "Quản lý Modal thêm người dùng mới, xử lý thuộc tính lồng nhau bằng Spread Operator và gửi dữ liệu lên App."),
        ("ResultTable", "Child (Con)", "users (array), loading (bool), error, editing (object | null)", "Gọi API JSONPlaceholder bằng useEffect; lắng nghe user mới; lọc theo keyword; mở Modal Sửa (Deep Copy) và Xóa.")
    ]
    add_styled_table(doc, comp_headers, comp_rows, [Inches(1.1), Inches(0.95), Inches(1.85), Inches(2.7)], center_cols=[0, 1])

    # ==================== PHẦN 2: SƠ ĐỒ LUỒNG DỮ LIỆU & STATE LIFTING ====================
    add_heading_1(doc, "2. SƠ ĐỒ LUỒNG DỮ LIỆU (STATE LIFTING) & GIẢI THÍCH THIẾT KẾ")
    add_body_p(
        doc,
        "Mô hình luồng dữ liệu của ứng dụng tuân thủ nghiêm ngặt nguyên lý Luồng dữ liệu một chiều (One-Way Data Flow) kết hợp kỹ thuật Nâng State lên component cha (State Lifting) cho hai chức năng chính: Tìm kiếm và Thêm người dùng."
    )

    pic_p2 = doc.add_paragraph()
    pic_p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pic_p2.paragraph_format.space_after = Pt(4)
    pic_p2.add_run().add_picture(flow_img_path, width=Inches(6.5))

    cap2 = doc.add_paragraph()
    cap2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap2.paragraph_format.space_after = Pt(8)
    rc2 = cap2.add_run("Hình 2.1: Sơ đồ luồng dữ liệu hai nhánh (State Lifting) giữa SearchForm, AddUser, App và ResultTable.")
    rc2.font.name = "Arial"
    rc2.font.size = Pt(9.5)
    rc2.font.italic = True
    rc2.font.color.rgb = RGBColor(100, 116, 139)

    add_heading_2(doc, "2.1. Chi tiết hai luồng dữ liệu chính")
    add_body_p(
        doc,
        "SearchForm (người dùng nhập text) -> gọi prop callback onChangeValue(e.target.value) -> cập nhật state kw tại App -> App re-render và truyền prop keyword={kw} xuống ResultTable -> ResultTable thực thi users.filter(...) để hiển thị danh sách khớp từ khóa.",
        bold_prefix="• Luồng Tìm kiếm (SearchForm -> App -> ResultTable): "
    )
    add_body_p(
        doc,
        "AddUser (người dùng điền form và bấm 'Thêm') -> gọi prop callback onAdd(user) -> cập nhật state newUser tại App -> App truyền prop user={newUser} xuống ResultTable -> Hook useEffect([user]) trong ResultTable nối user mới vào state users[] -> ResultTable gọi onAdded() để đưa newUser ở App trở về null.",
        bold_prefix="• Luồng Thêm người dùng (AddUser -> App -> ResultTable -> App): "
    )

    add_heading_2(doc, "2.2. Giải thích tại sao state kw và newUser đặt ở App mà không đặt ở SearchForm hay AddUser?")
    add_body_p(
        doc,
        "React hoạt động theo cơ chế Luồng dữ liệu một chiều (One-Way Data Flow): dữ liệu (props) chỉ có thể truyền từ Component Cha xuống Component Con. Hai component ngang hàng (Sibling Components) nằm chung một cấp dưới App là SearchForm với ResultTable, hoặc AddUser với ResultTable KHÔNG có đường truyền trực tiếp state/props sang cho nhau.",
        bold_prefix="1) Giới hạn truyền dữ liệu giữa các Sibling Components: "
    )
    add_body_p(
        doc,
        "State kw sinh ra từ thao tác nhập liệu ở SearchForm, nhưng nơi tiêu thụ và sử dụng kw để lọc danh sách (users.filter) lại là ResultTable. Nếu đặt kw ở bên trong SearchForm, ResultTable sẽ bị cô lập và không thể đọc được giá trị từ khóa mà người dùng đang tìm.",
        bold_prefix="2) Nếu đặt kw ở SearchForm: "
    )
    add_body_p(
        doc,
        "State newUser được tạo ra tại form của AddUser, nhưng mảng danh sách users gốc lại nằm trong ResultTable. Nếu đặt newUser ở bên trong AddUser, ResultTable sẽ không thể biết khi nào có người dùng mới được bấm 'Thêm' để nối vào bảng dữ liệu.",
        bold_prefix="3) Nếu đặt newUser ở AddUser: "
    )
    add_body_p(
        doc,
        "Bằng cách nâng cả kw và newUser lên Component cha chung gần nhất là App (State Lifting), App đóng vai trò làm 'Single Source of Truth' (Nguồn chân lý duy nhất). App cung cấp hàm cập nhật (setter) cho component phát sinh sự kiện (SearchForm, AddUser) và truyền giá trị state mới nhất xuống component tiêu thụ (ResultTable). Ngược lại, các state chỉ phục vụ giao diện cục bộ (như adding trong AddUser hay editing trong ResultTable) không cần chia sẻ nên vẫn được giữ ở component con để tối ưu phạm vi quản lý.",
        bold_prefix="4) Giải pháp State Lifting tại App: "
    )

    # ==================== PHẦN 3: MÔ TẢ PROPS ====================
    add_heading_1(doc, "3. BẢNG MÔ TẢ PROPS TRUYỀN TỪ APP XUỐNG CÁC COMPONENT CON")
    add_body_p(
        doc,
        "Bảng dưới đây tổng hợp toàn bộ 5 Props được component App truyền xuống cho 3 component con SearchForm, AddUser và ResultTable:"
    )

    prop_headers = ["Component nhận", "Tên Prop", "Giá trị truyền từ App", "Kiểu dữ liệu", "Chiều dữ liệu", "Mô tả vai trò & cơ chế hoạt động"]
    prop_rows = [
        ("SearchForm", "onChangeValue", "setKeyword", "Function", "Con -> Cha\n(Callback)", "Truyền hàm cập nhật state kw. Khi input xảy ra sự kiện onChange, SearchForm gọi onChangeValue(e.target.value) để cập nhật kw trên App."),
        ("AddUser", "onAdd", "setNewUser", "Function", "Con -> Cha\n(Callback)", "Truyền hàm cập nhật state newUser. Khi form hợp lệ và người dùng bấm Thêm, AddUser gọi onAdd(user) để gửi object user mới lên App."),
        ("ResultTable", "keyword", "kw", "String", "Cha -> Con\n(Data)", "Truyền giá trị từ khóa tìm kiếm hiện tại từ App xuống để ResultTable lọc mảng users theo name và username (không phân biệt hoa/thường)."),
        ("ResultTable", "user", "newUser", "Object | null", "Cha -> Con\n(Data)", "Truyền thông tin người dùng mới từ App xuống. Hook useEffect([user]) theo dõi prop này; khi khác null sẽ thêm vào đầu/cuối danh sách users."),
        ("ResultTable", "onAdded", "() => setNewUser(null)", "Function", "Con -> Cha\n(Callback)", "Callback để ResultTable báo lại cho App sau khi đã thêm xong user vào bảng, reset state newUser về null nhằm tránh thêm lặp ở lần render sau.")
    ]
    add_styled_table(doc, prop_headers, prop_rows, [Inches(0.95), Inches(1.0), Inches(1.15), Inches(0.8), Inches(0.85), Inches(1.85)], center_cols=[0, 3, 4])

    # ==================== PHẦN 4: LỖI / WARNING TRONG QUÁ TRÌNH CODE ====================
    add_heading_1(doc, "4. CÁC LỖI (ERROR / WARNING) GẶP PHẢI, NGUYÊN NHÂN VÀ CÁCH SỬA")
    add_body_p(
        doc,
        "Trong quá trình xây dựng và kiểm thử ứng dụng theo tài liệu hướng dẫn, em đã gặp và xử lý 3 lỗi/cảnh báo quan trọng liên quan đến cơ chế render của React, tham chiếu đối tượng lồng nhau (nested object) và định danh key trong danh sách:"
    )

    # Lỗi 1
    add_heading_2(doc, "4.1. Lỗi 1 (Runtime Error): Too many re-renders khi bấm mở Modal Thêm người dùng")
    add_body_p(doc, "Ứng dụng bị đứng/trắng trang khi component AddUser render Modal, Console trình duyệt báo lỗi đỏ:", bold_prefix="• Hiện tượng / Thông báo lỗi: ")
    add_code_block(doc, "Uncaught Error: Too many re-renders. React limits the number of renders to prevent an infinite loop.\n    at renderWithHooks (react-dom.development.js:...)\n    at AddUser (<anonymous>:...)")
    add_body_p(
        doc,
        "Trong đoạn mã mẫu Modal của AddUser ở tài liệu hướng dẫn (Bước 5, trang 6), nút Hủy được viết là <button onClick={setAdding(false)}>Hủy</button>. Việc viết cặp dấu ngoặc đơn () khiến JavaScript thực thi hàm setAdding(false) NGAY LẬP TỨC trong lúc component đang render, thay vì truyền vào một hàm xử lý sự kiện (event handler). Lệnh setAdding thay đổi state ngay trong quá trình render, ép React phải render lại liên tục tạo thành vòng lặp vô tận.",
        bold_prefix="• Nguyên nhân: "
    )
    add_body_p(
        doc,
        "Bọc lời gọi setAdding(false) vào trong một Arrow Function (hoặc tách thành hàm handleCancel) để React chỉ thực thi khi người dùng thực sự click vào nút Hủy:",
        bold_prefix="• Cách khắc phục: "
    )
    add_code_block(doc, """// Code lỗi:
<button onClick={setAdding(false)}>Hủy</button>

// Code đã sửa:
<button className="btn-secondary" onClick={() => setAdding(false)}>Hủy</button>""")

    # Lỗi 2
    add_heading_2(doc, "4.2. Lỗi 2 (Logic & State Mutation): Sửa địa chỉ trên Modal làm thay đổi dữ liệu gốc khi chưa Lưu")
    add_body_p(
        doc,
        "Khi nhấn nút 'Sửa' một người dùng ở ResultTable và gõ thay đổi các trường Street, Suite hoặc City trên form Modal, dữ liệu của dòng người dùng tương ứng trong bảng phía sau lập tức bị thay đổi theo. Nếu nhấn 'Hủy', bảng vẫn giữ giá trị sai vừa gõ thay vì khôi phục dữ liệu cũ.",
        bold_prefix="• Hiện tượng sai: "
    )
    add_body_p(
        doc,
        "Nếu sao chép đối tượng sang state editing bằng Shallow Copy (setEditing({ ...user })), toán tử spread chỉ sao chép các thuộc tính cấp 1, còn thuộc tính lồng nhau editing.address vẫn trỏ tới CÙNG một vùng nhớ (tham chiếu) với user.address trong state users. Do đó mọi thay đổi lên address ở form Sửa đều làm biến đổi (mutate) trực tiếp state gốc của bảng.",
        bold_prefix="• Nguyên nhân: "
    )
    add_body_p(
        doc,
        "Áp dụng kỹ thuật Deep Copy ở cả hàm editUser() lẫn hàm cập nhật handleEditChange() bằng Spread Operator 2 tầng:",
        bold_prefix="• Cách khắc phục: "
    )
    add_code_block(doc, """// 1. Deep Copy khi mở Modal Sửa:
function editUser(targetUser) {
    setEditing({
        ...targetUser,
        address: { ...(targetUser.address || { street: "", suite: "", city: "" }) }
    });
}
// 2. Tạo bản sao bất biến cho address khi nhập liệu trong form Sửa:
function handleEditChange(field, value) {
    if (["street", "suite", "city"].includes(field)) {
        setEditing((prev) => ({ ...prev, address: { ...prev.address, [field]: value } }));
    } else {
        setEditing((prev) => ({ ...prev, [field]: value }));
    }
}""")

    # Lỗi 3
    add_heading_2(doc, "4.3. Lỗi 3 (React Console Warning): Trùng lặp key khi Xóa rồi Thêm người dùng mới")
    add_body_p(doc, "Sau khi xóa một người dùng bất kỳ ở giữa bảng rồi thực hiện Thêm người dùng mới, Console của React phát ra cảnh báo:", bold_prefix="• Hiện tượng / Thông báo Warning: ")
    add_code_block(doc, "Warning: Encountered two children with the same key, `10`. Keys should be unique so that components maintain their identity across updates.")
    add_body_p(
        doc,
        "Nếu gán id mới bằng công thức prev.length + 1 (như gợi ý cơ bản ở trang 6): danh sách ban đầu có 10 user (id từ 1 đến 10). Khi xóa 1 user (ví dụ id=3), mảng còn 9 phần tử nhưng vẫn chứa phần tử có id=10. Khi thêm user mới, prev.length + 1 cho kết quả 9 + 1 = 10, dẫn đến 2 dòng <tr key={u.id}> có cùng key='10'.",
        bold_prefix="• Nguyên nhân: "
    )
    add_body_p(
        doc,
        "Tính id mới dựa trên giá trị id lớn nhất hiện có trong mảng thay vì dùng prev.length + 1, kết hợp truyền đầy đủ key={u.id} trên thẻ <tr>:",
        bold_prefix="• Cách khắc phục: "
    )
    add_code_block(doc, """const nextId = prev.length > 0 ? Math.max(...prev.map((u) => u.id)) + 1 : 1;
return [...prev, { ...user, id: nextId }];""")

    # ==================== PHẦN 5: TRIỂN KHAI LÊN HOST & KIỂM THỬ ====================
    add_heading_1(doc, "5. KẾT QUẢ UPLOAD ỨNG DỤNG LÊN HOST, KIỂM THỬ & MÃ NGUỒN")
    add_body_p(
        doc,
        "Ứng dụng hoàn chỉnh đã được đẩy (push) lên GitHub Repository của môn học và triển khai công khai trên dịch vụ hosting tĩnh GitHub Pages, đồng thời tích hợp liên kết trực tiếp vào menu điều hướng của trang chủ dự án."
    )

    host_headers = ["Thông tin triển khai", "Đường dẫn (URL) / Trạng thái"]
    host_rows = [
        ("Link ứng dụng React CRUD (Live Host)", "https://vuhoabinh2416137.github.io/IT4409/Bai%20tap%20React%20-%20CRUD/index.html"),
        ("Link Trang chủ tổng hợp các bài tập", "https://vuhoabinh2416137.github.io/IT4409/"),
        ("Link Kho mã nguồn (GitHub Repository)", "https://github.com/vuhoabinh2416137/IT4409"),
        ("Đường dẫn file trong Repository", "Bai tap React - CRUD/index.html"),
        ("Thư viện & Công nghệ sử dụng", "React 18 (CDN UMD), ReactDOM 18, Babel Standalone, Fetch API (JSONPlaceholder)")
    ]
    add_styled_table(doc, host_headers, host_rows, [Inches(2.3), Inches(4.3)])

    add_heading_2(doc, "5.1. Bảng kiểm thử đầy đủ các chức năng CRUD & Tìm kiếm")
    crud_headers = ["Chức năng CRUD", "Thao tác kiểm thử", "Kết quả thực tế"]
    crud_rows = [
        ("Read (Đọc & Hiển thị)", "Tải trang lần đầu", "Gọi API JSONPlaceholder thành công, hiển thị đầy đủ 10 người dùng (ID, Họ tên, Username, Email, Địa chỉ, SĐT, Website)."),
        ("Search (Tìm kiếm)", "Gõ từ khóa vào SearchForm", "Lọc danh sách theo thời gian thực (không phân biệt hoa thường) theo Name và Username; cập nhật bộ đếm hiển thị."),
        ("Create (Thêm mới)", "Bấm '+ Thêm người dùng', nhập đủ trường & Lưu", "Validate Name/Username bắt buộc; thêm thành công bản ghi với ID duy nhất vào bảng và đóng Modal."),
        ("Update (Chỉnh sửa)", "Bấm 'Sửa', thay đổi thông tin kể cả Address", "Hiển thị Modal chứa dữ liệu Deep Copy; khi chưa lưu hoặc bấm Hủy không ảnh hưởng dữ liệu gốc; bấm Lưu cập nhật đúng dòng."),
        ("Delete (Xóa)", "Bấm 'Xóa' trên một dòng bất kỳ", "Xóa trực tiếp người dùng khỏi state mảng bằng filter() và cập nhật lại giao diện bảng ngay lập tức.")
    ]
    add_styled_table(doc, crud_headers, crud_rows, [Inches(1.4), Inches(2.0), Inches(3.2)])

    add_heading_2(doc, "5.2. Mã nguồn hoàn chỉnh phần React (Bai tap React - CRUD/index.html)")
    index_path = os.path.join(CURRENT_DIR, "index.html")
    with open(index_path, "r", encoding="utf-8") as f:
        full_html = f.read()
    start_marker = '<script type="text/babel">'
    end_marker = '</script>'
    start_idx = full_html.find(start_marker)
    end_idx = full_html.rfind(end_marker)
    if start_idx != -1 and end_idx != -1:
        react_code = full_html[start_idx + len(start_marker):end_idx].strip()
    else:
        react_code = full_html.strip()
    add_code_block(doc, react_code)

    # Save .docx and .doc
    docx_path = os.path.join(CURRENT_DIR, "Bao_cao_Bai_tap_React_CRUD.docx")
    doc_path = os.path.join(CURRENT_DIR, "Bao_cao_Bai_tap_React_CRUD.doc")

    doc.save(docx_path)
    shutil.copyfile(docx_path, doc_path)

    print(f"Generated diagrams:\n  - {tree_img_path}\n  - {flow_img_path}")
    print(f"Generated report:\n  - {docx_path}\n  - {doc_path}")


if __name__ == "__main__":
    build_report()
