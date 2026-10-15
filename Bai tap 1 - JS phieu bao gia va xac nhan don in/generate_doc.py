import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import os

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_hex)
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

doc = docx.Document()

# Set standard margins (1 inch)
for section in doc.sections:
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)

# Title
title_p = doc.add_paragraph()
title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
title_run = title_p.add_run("BÁO CÁO BÀI TẬP 1: JAVASCRIPT")
title_run.font.name = "Arial"
title_run.font.size = Pt(20)
title_run.font.bold = True
title_run.font.color.rgb = RGBColor(31, 78, 121)

sub_p = doc.add_paragraph()
sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
sub_run = sub_p.add_run("Phiếu Báo Giá & Phiếu Xác Nhận Đơn In (Blakletterpress)")
sub_run.font.name = "Arial"
sub_run.font.size = Pt(13)
sub_run.font.italic = True
sub_run.font.color.rgb = RGBColor(89, 89, 89)

doc.add_paragraph() # Spacer

# ==================== PHẦN 1: BÀI 1A ====================
h1 = doc.add_paragraph()
r = h1.add_run("1. BÀI 1A — ĐIỀN VÀO CHỖ TRỐNG (js/bai-a-bao-gia.js)")
r.font.name = "Arial"
r.font.size = Pt(14)
r.font.bold = True
r.font.color.rgb = RGBColor(31, 78, 121)

p_intro_1a = doc.add_paragraph()
p_intro_1a.add_run("Dưới đây là bảng tổng hợp chi tiết đáp án cho 19 chỗ trống trong file ")
r_code = p_intro_1a.add_run("js/bai-a-bao-gia.js")
r_code.font.bold = True
p_intro_1a.add_run(", bao gồm từ khóa / toán tử / ký hiệu trong code, kết quả dự đoán in ra console, và giải thích chi tiết cơ chế hoạt động của JavaScript:")

# Table for 19 blanks
table_1a = doc.add_table(rows=1, cols=4)
table_1a.alignment = WD_TABLE_ALIGNMENT.CENTER
table_1a.autofit = False

headers = ["STT", "Đáp án điền", "Loại", "Giải thích chi tiết & cơ chế JS"]
col_widths = [Inches(0.6), Inches(1.8), Inches(0.9), Inches(3.2)]

hdr_cells = table_1a.rows[0].cells
for i, title in enumerate(headers):
    hdr_cells[i].text = title
    set_cell_background(hdr_cells[i], "1F4E79")
    set_cell_margins(hdr_cells[i], top=120, bottom=120, left=120, right=120)
    p = hdr_cells[i].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in p.runs:
        run.font.name = "Arial"
        run.font.size = Pt(10)
        run.font.bold = True
        run.font.color.rgb = RGBColor(255, 255, 255)

data_1a = [
    ("(1)", "const", "Code", "Khai báo hằng số TEN_XUONG không bao giờ được gán lại."),
    ("(2)", "let", "Code", "Khai báo biến soLuong có phạm vi khối (block scope) và cho phép gán lại giá trị ở Phần 2."),
    ("(3)", "300200", "Dự đoán", "Toán tử '+' giữa number (300) và string ('200') thực hiện ép kiểu số sang chuỗi và nối chuỗi thành '300200'."),
    ("(4)", "true", "Dự đoán", "Toán tử so sánh bằng lỏng lẻo '==' tự động ép kiểu chuỗi '200' thành số 200 trước khi so sánh nên trả về true."),
    ("(5)", "false", "Dự đoán", "Toán tử so sánh nghiêm ngặt '===' không ép kiểu; hai vế khác kiểu dữ liệu (string vs number) nên trả về false."),
    ("(6)", "let", "Code", "Biến giamGia cần được gán lại bên trong khối if và tiếp tục sử dụng hợp lệ sau khối if (phạm vi bao ngoài)."),
    ("(7)", "ReferenceError (lyDo is not defined)", "Dự đoán", "Biến lyDo khai báo bằng const bên trong khối if có phạm vi khối (block-scoped), bên ngoài khối if không tồn tại nên quăng lỗi ReferenceError."),
    ("(8)", "undefined", "Dự đoán", "Cơ chế Hoisting của var đưa khai báo maDon lên đầu scope nhưng khởi tạo giá trị ban đầu là undefined."),
    ("(9)", "báo lỗi ReferenceError (Cannot access 'maDon' before initialization)", "Dự đoán", "Biến khai báo bằng let được hoisted nhưng nằm trong Vùng chết tạm thời (TDZ - Temporal Dead Zone) cho đến khi dòng khai báo được thực thi; truy cập sớm sẽ gây lỗi ReferenceError."),
    ("(10)", "==", "Code", "Trong JS, phép so sánh 'ghiChu == null' đúng khi ghiChu là null HOẶC undefined, và sai với '' hoặc 0 (thỏa mãn chỉ dùng 1 phép so sánh)."),
    ("(11)", "500", "Dự đoán", "Object là kiểu dữ liệu tham chiếu (reference type). donKhach = donMau trỏ chung địa chỉ ô nhớ, do đó cập nhật donKhach.soLuong = 500 cũng làm donMau.soLuong thành 500."),
    ("(12)", "true", "Dự đoán", "Toán tử '===' khi so sánh 2 object sẽ so sánh địa chỉ tham chiếu; vì donMau và donKhach cùng trỏ tới một object duy nhất nên trả về true."),
    ("(13)", "(B)", "Dự đoán", "donMau khai báo bằng const nên không thể gán lại cho object mới (B: donMau = {} sẽ quăng lỗi TypeError). Ngược lại, sửa thuộc tính (A: donMau.giay = 'Kraft') hoàn toàn hợp lệ."),
    ("(14)", "{", "Code", "Mở ngoặc nhọn '{' phải nằm cùng dòng với return để tránh cơ chế ASI (Automatic Semicolon Insertion) tự chèn dấu ';' sau return."),
    ("(15)", "undefined", "Dự đoán", "Nếu return đứng một mình một dòng, ASI sẽ tự động chèn dấu ';' thành 'return;', khiến hàm dừng ngay và trả về undefined."),
    ("(16)", "` (dấu backtick)", "Code", "Template literal trong JavaScript bắt buộc mở và đóng bằng dấu backtick (chuỗi đa dòng và nội suy biến)."),
    ("(17)", "${soLuong}", "Code", "Cú pháp nội suy biến ${soLuong} bên trong template literal để hiển thị số lượng (500)."),
    ("(18)", "log", "Code", "Phương thức console.log(phieu) dùng để in thông tin thông thường ra Console."),
    ("(19)", "table", "Code", "Phương thức console.table(...) dùng để in mảng các object trực quan dưới dạng bảng trong Console.")
]

for row_idx, data in enumerate(data_1a):
    row_cells = table_1a.add_row().cells
    bg_color = "F2F2F2" if row_idx % 2 == 1 else "FFFFFF"
    for col_idx, text in enumerate(data):
        row_cells[col_idx].text = text
        set_cell_background(row_cells[col_idx], bg_color)
        set_cell_margins(row_cells[col_idx], top=80, bottom=80, left=100, right=100)
        p = row_cells[col_idx].paragraphs[0]
        if col_idx in [0, 2]:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in p.runs:
            run.font.name = "Arial"
            run.font.size = Pt(9.5)
            if col_idx == 1:
                run.font.bold = True

for row in table_1a.rows:
    for i, w in enumerate(col_widths):
        row.cells[i].width = w

doc.add_paragraph() # Spacer

# Kết quả in ra trang index.html
h2_phieu = doc.add_paragraph()
r = h2_phieu.add_run("Kết quả phiếu báo giá hiển thị trên giao diện:")
r.font.name = "Arial"
r.font.size = Pt(11)
r.font.bold = True

res_phieu = """=== Blakletterpress ===
Số lượng : 500 thiệp
Tạm tính : 600000 đ
Giảm giá : 60000 đ (10 %)
VAT      : 54000 đ
Tổng     : 594000 đ
Ghi chú  : (không có)"""

p_phieu = doc.add_paragraph()
r_phieu = p_phieu.add_run(res_phieu)
r_phieu.font.name = "Consolas"
r_phieu.font.size = Pt(9.5)
r_phieu.font.bold = True
r_phieu.font.color.rgb = RGBColor(0, 102, 0)

# Mã nguồn hoàn chỉnh bài 1a
h2_1a = doc.add_paragraph()
r = h2_1a.add_run("Mã nguồn hoàn chỉnh js/bai-a-bao-gia.js:")
r.font.name = "Arial"
r.font.size = Pt(11)
r.font.bold = True

current_dir = os.path.dirname(os.path.abspath(__file__))
with open(os.path.join(current_dir, "js", "bai-a-bao-gia.js"), "r", encoding="utf-8") as f:
    code_content_1a = f.read()

code_p1a = doc.add_paragraph()
r_code_1a = code_p1a.add_run(code_content_1a)
r_code_1a.font.name = "Consolas"
r_code_1a.font.size = Pt(8.5)
r_code_1a.font.color.rgb = RGBColor(40, 40, 40)

doc.add_page_break()

# ==================== PHẦN 2: BÀI 1B ====================
h1_b = doc.add_paragraph()
r = h1_b.add_run("2. BÀI 1B — VIẾT CODE XÁC NHẬN ĐƠN IN (js/bai-b-phieu-in.js)")
r.font.name = "Arial"
r.font.size = Pt(14)
r.font.bold = True
r.font.color.rgb = RGBColor(31, 78, 121)

p_intro_1b = doc.add_paragraph()
p_intro_1b.add_run("Bài 1b yêu cầu cài đặt 4 hàm trong ")
r_code = p_intro_1b.add_run("js/bai-b-phieu-in.js")
r_code.font.bold = True
p_intro_1b.add_run(" để vượt qua toàn bộ 12 testcase tự động (T1 đến T12) của bộ kiểm tra ")
r_kt = p_intro_1b.add_run("js/kiem-tra.js")
r_kt.font.bold = True
p_intro_1b.add_run(":")

# 4 functions explanation
funcs = [
    ("1. Hàm tinhTong(soLuong, donGia, phiGiao)", 
     "• Mục đích: Tính tổng tiền gồm tạm tính (soLuong × donGia) + thuế VAT (VAT_PHAN_TRAM %) + phí giao hàng phiGiao.\n"
     "• Xử lý kiểu dữ liệu: Sử dụng Number() để chuyển đổi các tham số nếu chúng được truyền vào dạng chuỗi từ form.\n"
     "• Trả về kiểu number: Đảm bảo phép toán cộng không bị nối chuỗi.",
     """function tinhTong(soLuong, donGia, phiGiao) {
	const sl = Number(soLuong);
	const dg = Number(donGia);
	const pg = Number(phiGiao);
	const tamTinh = sl * dg;
	const vat = tamTinh * VAT_PHAN_TRAM / 100;
	return tamTinh + vat + pg;
}"""),
    
    ("2. Hàm moTaInLai(soLan)",
     "• Mục đích: Diễn giải trạng thái in ấn dựa vào số lần in.\n"
     "• Xử lý null / undefined: Dùng soLan == null để bắt cả null và undefined, trả về 'Chưa rõ'.\n"
     "• Xử lý in lần đầu: Nếu Number(soLan) === 0, trả về 'In lần đầu'.\n"
     "• Xử lý in lại lần n: Với n dương, dùng template literal trả về `In lại lần ${soLan}`.",
     """function moTaInLai(soLan) {
	if (soLan == null) {
		return "Chưa rõ";
	}
	if (Number(soLan) === 0) {
		return "In lần đầu";
	}
	return `In lại lần ${soLan}`;
}"""),

    ("3. Hàm inPhieu(don)",
     "• Mục đích: Trả về chuỗi phiếu xác nhận gồm chính xác 5 dòng định dạng bằng template literal, không có dòng hay khoảng trắng thừa.\n"
     "• Dòng 1: `${TEN_XUONG.toUpperCase()}` chuyển tên xưởng thành chữ in hoa (BLAKLETTERPRESS).\n"
     "• Dòng 4 & 5: Tái sử dụng tinhTong(...) và moTaInLai(don.soLanIn).",
     """function inPhieu(don) {
	return `=== ${TEN_XUONG.toUpperCase()} ===
Mã đơn : ${don.ma}
Khách  : ${don.khach.ten} (${don.khach.sdt})
Tổng   : ${tinhTong(don.soLuong, don.donGia, don.phiGiao)} đ
Ghi chú: ${moTaInLai(don.soLanIn)}`;
}"""),

    ("4. Hàm taoDonInLai(donCu, sdtMoi)",
     "• Mục đích: Tạo một đối tượng đơn hàng mới từ đơn cũ với số điện thoại khách mới và số lần in tăng thêm 1.\n"
     "• Đảm bảo tính bất biến (Immutability & Deep Clone): Sử dụng object spread {...donCu} và spread lồng cho object khach {...donCu.khach, sdt: sdtMoi}. Nhờ đó đối tượng con khach cũng được nhân bản độc lập, việc thay đổi thuộc tính trên đơn mới (như khach.ten) hoàn toàn không ảnh hưởng đến đơn cũ (thỏa mãn triệt để testcase T11 và T12).",
     """function taoDonInLai(donCu, sdtMoi) {
	return {
		...donCu,
		khach: {
			...donCu.khach,
			sdt: sdtMoi
		},
		soLanIn: Number(donCu.soLanIn) + 1
	};
}""")
]

for title, desc, code in funcs:
    hp = doc.add_paragraph()
    r = hp.add_run(title)
    r.font.name = "Arial"
    r.font.size = Pt(11)
    r.font.bold = True
    r.font.color.rgb = RGBColor(31, 78, 121)

    dp = doc.add_paragraph()
    r_desc = dp.add_run(desc)
    r_desc.font.name = "Arial"
    r_desc.font.size = Pt(9.5)

    cp = doc.add_paragraph()
    r_c = cp.add_run(code)
    r_c.font.name = "Consolas"
    r_c.font.size = Pt(9)
    r_c.font.color.rgb = RGBColor(30, 30, 30)

doc.add_paragraph() # Spacer

# Bảng kết quả 12 testcases
h2_tc = doc.add_paragraph()
r = h2_tc.add_run("Bảng kết quả chạy bộ kiểm tra tự động js/kiem-tra.js (12/12 Đạt):")
r.font.name = "Arial"
r.font.size = Pt(11)
r.font.bold = True

table_tc = doc.add_table(rows=1, cols=3)
table_tc.alignment = WD_TABLE_ALIGNMENT.CENTER
table_tc.autofit = False

tc_headers = ["Mã Testcase", "Nội dung kiểm tra", "Kết quả"]
tc_col_widths = [Inches(1.2), Inches(4.3), Inches(1.1)]

hdr_cells_tc = table_tc.rows[0].cells
for i, title in enumerate(tc_headers):
    hdr_cells_tc[i].text = title
    set_cell_background(hdr_cells_tc[i], "1F4E79")
    set_cell_margins(hdr_cells_tc[i], top=100, bottom=100, left=100, right=100)
    p = hdr_cells_tc[i].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in p.runs:
        run.font.name = "Arial"
        run.font.size = Pt(10)
        run.font.bold = True
        run.font.color.rgb = RGBColor(255, 255, 255)

testcases_data = [
    ("T1", "tinhTong nhận số: (250, 1500, 30000) → 442500", "Đạt"),
    ("T2", "tinhTong nhận chuỗi: (\"250\", \"1500\", \"30000\") → 442500", "Đạt"),
    ("T3", "tinhTong trả về kiểu number", "Đạt"),
    ("T4", "tinhTong(\"100\", \"2000\", \"0\") → 220000", "Đạt"),
    ("T5", "moTaInLai(0) → \"In lần đầu\"", "Đạt"),
    ("T6", "moTaInLai(2) → \"In lại lần 2\"", "Đạt"),
    ("T7", "moTaInLai(null) và moTaInLai(undefined) → \"Chưa rõ\"", "Đạt"),
    ("T8", "inPhieu đúng từng ký tự (5 dòng)", "Đạt"),
    ("T9", "taoDonInLai trả về object MỚI (khác đơn cũ)", "Đạt"),
    ("T10", "đơn mới: sdt mới, soLanIn tăng 1", "Đạt"),
    ("T11", "đơn cũ KHÔNG đổi sdt và soLanIn", "Đạt"),
    ("T12", "sửa tên khách trên đơn mới không ảnh hưởng đơn cũ", "Đạt")
]

for row_idx, data in enumerate(testcases_data):
    row_cells = table_tc.add_row().cells
    bg_color = "E2EFDA" if row_idx % 2 == 1 else "FFFFFF"
    for col_idx, text in enumerate(data):
        row_cells[col_idx].text = text
        set_cell_background(row_cells[col_idx], bg_color)
        set_cell_margins(row_cells[col_idx], top=60, bottom=60, left=80, right=80)
        p = row_cells[col_idx].paragraphs[0]
        if col_idx in [0, 2]:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in p.runs:
            run.font.name = "Arial"
            run.font.size = Pt(9.5)
            if col_idx == 2:
                run.font.bold = True
                run.font.color.rgb = RGBColor(56, 86, 35)

for row in table_tc.rows:
    for i, w in enumerate(tc_col_widths):
        row.cells[i].width = w

doc.add_paragraph() # Spacer

# Mã nguồn hoàn chỉnh bài 1b
h2_1b = doc.add_paragraph()
r = h2_1b.add_run("Mã nguồn hoàn chỉnh js/bai-b-phieu-in.js:")
r.font.name = "Arial"
r.font.size = Pt(11)
r.font.bold = True

with open(os.path.join(current_dir, "js", "bai-b-phieu-in.js"), "r", encoding="utf-8") as f:
    code_content_1b = f.read()

code_p1b = doc.add_paragraph()
r_code_1b = code_p1b.add_run(code_content_1b)
r_code_1b.font.name = "Consolas"
r_code_1b.font.size = Pt(8.5)
r_code_1b.font.color.rgb = RGBColor(40, 40, 40)

doc.add_page_break()

# ==================== PHẦN 3: PHÂN TÍCH THỨ TỰ NẠP SCRIPT ====================
h1_c = doc.add_paragraph()
r = h1_c.add_run("3. PHÂN TÍCH THỨ TỰ NẠP FILE SCRIPT TRÊN CÁC TRANG HTML")
r.font.name = "Arial"
r.font.size = Pt(14)
r.font.bold = True
r.font.color.rgb = RGBColor(31, 78, 121)

p_script = doc.add_paragraph()
p_script.add_run("Trong mô hình nạp script đồng bộ chuẩn ở cuối thẻ <body> của HTML:\n\n")

p_script_idx = doc.add_paragraph()
r_idx_h = p_script_idx.add_run("• Trang index.html:\n")
r_idx_h.font.bold = True
p_script_idx.add_run("   <script src=\"js/ho-tro.js\"></script>\n"
                       "   <script src=\"js/bai-a-bao-gia.js\"></script>\n"
                       "-> Phân tích: File ho-tro.js được nạp trước để đăng ký hàm hienThi(noiDung) vào phạm vi toàn cục (global scope). Sau đó file bai-a-bao-gia.js được nạp vào mới có thể gọi hàm hienThi(phieu) thành công. Nếu đảo ngược thứ tự, trang web sẽ báo lỗi 'Uncaught ReferenceError: hienThi is not defined' khi chạy bai-a-bao-gia.js.\n\n")

r_ab_h = p_script_idx.add_run("• Trang about.html:\n")
r_ab_h.font.bold = True
p_script_idx.add_run("   <script src=\"js/ho-tro.js\"></script>\n"
                      "   <script src=\"js/bai-b-phieu-in.js\"></script>\n"
                      "   <script src=\"js/kiem-tra.js\"></script>\n"
                      "-> Phân tích:\n"
                      "   1. ho-tro.js được nạp đầu tiên để chuẩn bị sẵn hàm hienThi() phục vụ in văn bản ra giao diện.\n"
                      "   2. bai-b-phieu-in.js được nạp thứ hai nhằm khai báo các hằng số, biến đơn hàng và định nghĩa 4 hàm xử lý (tinhTong, moTaInLai, inPhieu, taoDonInLai), đồng thời gọi in phiếu ban đầu.\n"
                      "   3. kiem-tra.js được nạp sau cùng để thực thi bộ 12 testcases kiểm thử các hàm đã có trong bộ nhớ từ bước 2 và sử dụng hienThi() từ bước 1 để xuất bảng kết quả. Nếu kiem-tra.js được nạp trước bai-b-phieu-in.js, các lời gọi hàm trong testcases sẽ bị lỗi ReferenceError (chưa được định nghĩa).")

# Save document
output_path = os.path.join(current_dir, "Bao_cao_Bai_tap_1_JS.docx")
doc.save(output_path)
print(f"File created successfully at: {output_path}")
