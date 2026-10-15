import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import os
import shutil

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
title_run = title_p.add_run("BÁO CÁO BÀI TẬP 2: JAVASCRIPT")
title_run.font.name = "Arial"
title_run.font.size = Pt(20)
title_run.font.bold = True
title_run.font.color.rgb = RGBColor(31, 78, 121)

sub_p = doc.add_paragraph()
sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
sub_run = sub_p.add_run("Tin Tức Từ JSON & Giỏ Hàng Xưởng In (Blakletterpress)")
sub_run.font.name = "Arial"
sub_run.font.size = Pt(13)
sub_run.font.italic = True
sub_run.font.color.rgb = RGBColor(89, 89, 89)

doc.add_paragraph() # Spacer

# ==================== PHẦN 1: BÀI 2A ====================
h1 = doc.add_paragraph()
r = h1.add_run("1. BÀI 2A — ĐIỀN VÀO CHỖ TRỐNG (js/bai-a-tin-tuc.js)")
r.font.name = "Arial"
r.font.size = Pt(14)
r.font.bold = True
r.font.color.rgb = RGBColor(31, 78, 121)

p_intro_2a = doc.add_paragraph()
p_intro_2a.add_run("Dưới đây là bảng tổng hợp chi tiết đáp án cho 30 chỗ trống trong file ")
r_code = p_intro_2a.add_run("js/bai-a-tin-tuc.js")
r_code.font.bold = True
p_intro_2a.add_run(", bao gồm từ khóa, phương thức, biểu thức, kết quả dự đoán và giải thích chi tiết cơ chế hoạt động của JavaScript:")

# Table for 30 blanks
table_2a = doc.add_table(rows=1, cols=4)
table_2a.alignment = WD_TABLE_ALIGNMENT.CENTER
table_2a.autofit = False

headers = ["STT", "Đáp án điền", "Loại", "Giải thích chi tiết & cơ chế JS"]
col_widths = [Inches(0.6), Inches(1.8), Inches(0.9), Inches(3.2)]

hdr_cells = table_2a.rows[0].cells
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

data_2a = [
    ("(1)", "parse", "Code", "JSON.parse(duLieuTho) phân tích cú pháp chuỗi JSON thành mảng các object trong JavaScript."),
    ("(2)", "push", "Code", "Phương thức push() thêm một object vào CUỐI mảng dsBai."),
    ("(3)", "unshift", "Code", "Phương thức unshift() thêm một object vào ĐẦU mảng dsBai."),
    ("(4)", "pop", "Code", "Phương thức pop() lấy ra và xoá phần tử ở CUỐI mảng dsBai."),
    ("(5)", "6", "Dự đoán", "baiCuoi là phần tử vừa pop() khỏi cuối mảng (chính là object bài viết id = 6 vừa push ở dòng 21)."),
    ("(6)", "shift", "Code", "Phương thức shift() lấy ra và xoá phần tử ở ĐẦU mảng dsBai."),
    ("(7)", "5", "Dự đoán", "Mảng ban đầu có 5 bài; sau push (+1 = 6), unshift (+1 = 7), pop (-1 = 6), shift (-1 = 5), độ dài mảng quay về 5."),
    ("(8)", "[\"tac gia\"]", "Code", "Thêm thuộc tính có tên chứa khoảng trắng ('tac gia') bắt buộc dùng cú pháp ngoặc vuông baiMot[\"tac gia\"]."),
    ("(9)", "delete", "Code", "Toán tử delete xoá một thuộc tính khỏi object: delete baiMot.noiBat."),
    ("(10)", "undefined", "Dự đoán", "Thuộc tính noiBat đã bị xoá khỏi baiMot, khi truy cập vào thuộc tính không tồn tại sẽ trả về undefined."),
    ("(11)", "[GHIM] Mẫu thiệp cưới mùa thu", "Dự đoán", "Object là kiểu tham chiếu; baiGhim và dsBai[1] cùng trỏ tới một object trong bộ nhớ, nên thay đổi qua baiGhim làm dsBai[1].title đổi theo."),
    ("(12)", "filter", "Code", "Phương thức filter() lọc mảng, chỉ giữ lại các phần tử có draft === false (các bài đã xuất bản)."),
    ("(13)", "4", "Dự đoán", "Trong 5 bài gốc, chỉ có bài id 4 có draft: true; 4 bài còn lại (id 1, 2, 3, 5) có draft: false nên daDang.length = 4."),
    ("(14)", "map", "Code", "Phương thức map() biến đổi từng phần tử của mảng daDang thành một mảng mới chứa tiêu đề viết hoa."),
    ("(15)", "toUpperCase", "Code", "Phương thức toUpperCase() chuyển toàn bộ ký tự trong chuỗi tiêu đề thành chữ in hoa."),
    ("(16)", "bai.views", "Code", "Biểu thức cộng dồn lượt xem bai.views vào biến tích lũy tong trong mỗi vòng lặp của reduce()."),
    ("(17)", "0", "Code", "Giá trị khởi tạo (initialValue) bằng 0 cho biến tích lũy tong của hàm reduce()."),
    ("(18)", "2600", "Dự đoán", "Tổng lượt xem của 4 bài đã đăng: 320 + 870 + 540 + 870 = 2600."),
    ("(19)", "(B) một chuỗi bắt đầu bằng \"[object Object]\"", "Dự đoán", "Nếu bỏ initialValue 0, phần tử đầu tiên (object daDang[0]) sẽ làm giá trị khởi tạo, phép cộng object với number sẽ ép kiểu thành chuỗi '[object Object]870...'."),
    ("(20)", ">", "Code", "Dùng toán tử '>' để so sánh bai.views > max.views; khi bằng nhau thì biểu thức sai và giữ nguyên max (bài đứng trước)."),
    ("(21)", "2", "Dự đoán", "Hai bài có views cao nhất là id 2 và id 5 (cùng 870 lượt); do toán tử '>' không đổi khi bằng nhau nên giữ lại bài đứng trước là bài id 2."),
    ("(22)", "undefined", "Dự đoán", "Hàm mũi tên có thân khối lệnh { bai.views + 1; } nhưng không có từ khóa return nên trả về undefined; do đó thu[0] là undefined."),
    ("(23)", "of", "Code", "Vòng lặp for (const the of bai.tags) dùng để duyệt qua từng giá trị trong mảng thẻ tags."),
    ("(24)", "the", "Code", "Biến the đại diện cho tên thẻ, dùng làm khoá để đếm: demThe[the] = (demThe[the] || 0) + 1."),
    ("(25)", "2", "Dự đoán", "Thẻ 'mau moi' xuất hiện trong 2 bài viết (bài id 2 và bài id 5) nên demThe[\"mau moi\"] = 2."),
    ("(26)", "constructor", "Code", "Phương thức khởi tạo của class trong JavaScript ES6 được định nghĩa bằng từ khoá constructor."),
    ("(27)", "this", "Code", "Từ khóa this đại diện cho thể hiện (instance) của class BaiViet đang được khởi tạo."),
    ("(28)", "new", "Code", "Toán tử new dùng để tạo một đối tượng (instance) mới từ lớp BaiViet: new BaiViet(duLieu)."),
    ("(29)", "stringify", "Code", "JSON.stringify() chuyển đổi một object trong JavaScript thành chuỗi định dạng JSON."),
    ("(30)", "không", "Dự đoán", "JSON.stringify() chỉ tuần tự hoá các thuộc tính dữ liệu (data properties), tự động bỏ qua phương thức và hàm trên prototype (như tomTat).")
]

for row_idx, data in enumerate(data_2a):
    row_cells = table_2a.add_row().cells
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

for row in table_2a.rows:
    for i, w in enumerate(col_widths):
        row.cells[i].width = w

doc.add_paragraph() # Spacer

# Kết quả hiển thị trên trang news.html
h2_res = doc.add_paragraph()
r = h2_res.add_run("Kết quả thống kê hiển thị trên trang news.html:")
r.font.name = "Arial"
r.font.size = Pt(11)
r.font.bold = True

res_news = """Đã đăng: 4 bài — tổng 2600 lượt xem — bài hot: #2
Tiêu đề: KHAI TRƯƠNG XƯỞNG IN MỚI | [GHIM] MẪU THIỆP CƯỚI MÙA THU | HƯỚNG DẪN CHỌN GIẤY COTTON | ƯU ĐÃI IN DANH THIẾP
#su kien: 1 bài
#xuong in: 1 bài
#thiep cuoi: 1 bài
#mau moi: 2 bài
#huong dan: 1 bài
#giay: 1 bài
#khuyen mai: 1 bài"""

p_news = doc.add_paragraph()
r_n = p_news.add_run(res_news)
r_n.font.name = "Consolas"
r_n.font.size = Pt(9.5)
r_n.font.bold = True
r_n.font.color.rgb = RGBColor(0, 102, 0)

# Mã nguồn hoàn chỉnh bài 2a
h2_2a = doc.add_paragraph()
r = h2_2a.add_run("Mã nguồn hoàn chỉnh js/bai-a-tin-tuc.js:")
r.font.name = "Arial"
r.font.size = Pt(11)
r.font.bold = True

current_dir = os.path.dirname(os.path.abspath(__file__))
with open(os.path.join(current_dir, "js", "bai-a-tin-tuc.js"), "r", encoding="utf-8") as f:
    code_content_2a = f.read()

code_p2a = doc.add_paragraph()
r_code_2a = code_p2a.add_run(code_content_2a)
r_code_2a.font.name = "Consolas"
r_code_2a.font.size = Pt(8.5)
r_code_2a.font.color.rgb = RGBColor(40, 40, 40)

doc.add_page_break()

# ==================== PHẦN 2: BÀI 2B ====================
h1_b = doc.add_paragraph()
r = h1_b.add_run("2. BÀI 2B — VIẾT CODE GIỎ HÀNG XƯỞNG IN (js/bai-b-gio-hang.js)")
r.font.name = "Arial"
r.font.size = Pt(14)
r.font.bold = True
r.font.color.rgb = RGBColor(31, 78, 121)

p_intro_2b = doc.add_paragraph()
p_intro_2b.add_run("Bài 2b yêu cầu cài đặt class ")
r_code = p_intro_2b.add_run("GioHang")
r_code.font.bold = True
p_intro_2b.add_run(" và hàm ")
r_code2 = p_intro_2b.add_run("napGioHang(chuoiJSON)")
r_code2.font.bold = True
p_intro_2b.add_run(" trong file ")
r_code3 = p_intro_2b.add_run("js/bai-b-gio-hang.js")
r_code3.font.bold = True
p_intro_2b.add_run(" để vượt qua toàn bộ 20 testcases tự động (T1 đến T20) trong ")
r_kt = p_intro_2b.add_run("js/kiem-tra.js")
r_kt.font.bold = True
p_intro_2b.add_run(":")

# 8 methods/functions explanation
methods = [
    ("1. them(sanPham, soLuong)",
     "• Kiểm tra tính hợp lệ: soLuong phải là kiểu 'number', số nguyên (Number.isInteger) và > 0. Nếu không thoả mãn trả về false mà không làm thay đổi giỏ hàng.\n"
     "• Tính độc lập & bất biến: Dòng sản phẩm mới được tạo thành object độc lập { id, ten, loai, gia, soLuong }, hoàn toàn không giữ tham chiếu tới object sanPham bên ngoài hay DANH_MUC.\n"
     "• Cộng dồn số lượng: Nếu sản phẩm đã tồn tại trong giỏ (theo thuộc tính id), chỉ cộng dồn thuộc tính soLuong.",
     """them(sanPham, soLuong) {
	if (typeof soLuong !== "number" || !Number.isInteger(soLuong) || soLuong <= 0) {
		return false;
	}
	const dong = this.dsDong.find(d => d.id === sanPham.id);
	if (dong) {
		dong.soLuong += soLuong;
	} else {
		this.dsDong.push({
			id: sanPham.id,
			ten: sanPham.ten,
			loai: sanPham.loai,
			gia: sanPham.gia,
			soLuong: soLuong
		});
	}
	return true;
}"""),

    ("2. xoa(id)",
     "• Tìm kiếm vị trí dòng: Sử dụng findIndex() để xác định chỉ số phần tử có id cần xoá.\n"
     "• Thực hiện xoá: Dùng splice(idx, 1) để gỡ bỏ phần tử khỏi mảng this.dsDong và trả về true. Nếu không tìm thấy trả về false.",
     """xoa(id) {
	const idx = this.dsDong.findIndex(d => d.id === id);
	if (idx !== -1) {
		this.dsDong.splice(idx, 1);
		return true;
	}
	return false;
}"""),

    ("3. tongTien()",
     "• Tính tổng tiền của mọi dòng: Dùng reduce() nhân gia × soLuong của từng dòng và cộng dồn với giá trị ban đầu là 0. Giỏ rỗng trả về 0.",
     """tongTien() {
	return this.dsDong.reduce((tong, d) => tong + d.gia * d.soLuong, 0);
}"""),

    ("4. thongKeTheoLoai()",
     "• Thống kê doanh thu theo phân loại: Duyệt qua từng dòng trong giỏ hàng, gom nhóm tổng tiền theo thuộc tính loai. Giỏ rỗng trả về object rỗng {}.",
     """thongKeTheoLoai() {
	const tk = {};
	for (const d of this.dsDong) {
		tk[d.loai] = (tk[d.loai] || 0) + d.gia * d.soLuong;
	}
	return tk;
}"""),

    ("5. dongDatNhat()",
     "• Dòng có thành tiền lớn nhất: Nếu giỏ rỗng trả về null. Duyệt mảng để tìm dòng có gia × soLuong lớn nhất.\n"
     "• Xử lý bằng nhau: Sử dụng toán tử > để khi gặp 2 dòng bằng tiền thì luôn ưu tiên giữ dòng đứng trước.\n"
     "• Độc lập tham chiếu: Trả về một object bản sao hoàn toàn mới kèm thêm thuộc tính thanhTien.",
     """dongDatNhat() {
	if (this.dsDong.length === 0) return null;
	let maxDong = this.dsDong[0];
	let maxTien = maxDong.gia * maxDong.soLuong;
	for (let i = 1; i < this.dsDong.length; i++) {
		const tien = this.dsDong[i].gia * this.dsDong[i].soLuong;
		if (tien > maxTien) {
			maxTien = tien;
			maxDong = this.dsDong[i];
		}
	}
	return {
		id: maxDong.id,
		ten: maxDong.ten,
		loai: maxDong.loai,
		gia: maxDong.gia,
		soLuong: maxDong.soLuong,
		thanhTien: maxTien
	};
}"""),

    ("6. layDanhSach()",
     "• Bản sao an toàn: Sử dụng map() để tạo mảng mới và clone từng object bên trong, gán thêm thuộc tính thanhTien = gia × soLuong. Sửa mảng hoặc object từ layDanhSach() không ảnh hưởng tới dữ liệu giỏ hàng.",
     """layDanhSach() {
	return this.dsDong.map(d => ({
		id: d.id,
		ten: d.ten,
		loai: d.loai,
		gia: d.gia,
		soLuong: d.soLuong,
		thanhTien: d.gia * d.soLuong
	}));
}"""),

    ("7. xuatJSON()",
     "• Xuất dữ liệu lưu trữ: Trả về chuỗi JSON tuần tự hoá từ mảng this.dsDong, mỗi object đảm bảo đúng 5 thuộc tính (id, ten, loai, gia, soLuong), không chứa thanhTien.",
     """xuatJSON() {
	return JSON.stringify(this.dsDong);
}"""),

    ("8. napGioHang(chuoiJSON)",
     "• Khôi phục trạng thái giỏ: Nhận chuỗi JSON được tạo từ xuatJSON(), parse thành mảng và khởi tạo lại đối tượng GioHang mới có đầy đủ các phương thức hoạt động bình thường và độc lập.",
     """function napGioHang(chuoiJSON) {
	const g = new GioHang();
	try {
		const mang = JSON.parse(chuoiJSON);
		if (Array.isArray(mang)) {
			g.dsDong = mang.map(d => ({
				id: d.id,
				ten: d.ten,
				loai: d.loai,
				gia: d.gia,
				soLuong: d.soLuong
			}));
		}
	} catch (e) {
	}
	return g;
}""")
]

for title, desc, code in methods:
    hp = doc.add_paragraph()
    r = hp.add_run(title)
    r.font.name = "Arial"
    r.font.size = Pt(11)
    r.font.bold = True
    r.font.color.rgb = RGBColor(31, 78, 121)

    dp = doc.add_paragraph()
    r_dp = dp.add_run(desc)
    r_dp.font.name = "Arial"
    r_dp.font.size = Pt(9.5)

    cp = doc.add_paragraph()
    r_c = cp.add_run(code)
    r_c.font.name = "Consolas"
    r_c.font.size = Pt(8.5)
    r_c.font.color.rgb = RGBColor(40, 40, 40)

# Bảng kết quả 20 testcases
doc.add_paragraph()
h2_tc = doc.add_paragraph()
r = h2_tc.add_run("Bảng kết quả kiểm tra tự động 20 testcases (js/kiem-tra.js):")
r.font.name = "Arial"
r.font.size = Pt(11)
r.font.bold = True

table_tests = doc.add_table(rows=1, cols=3)
table_tests.alignment = WD_TABLE_ALIGNMENT.CENTER
table_tests.autofit = False

tc_headers = ["STT", "Nội dung kiểm tra", "Kết quả"]
tc_col_widths = [Inches(0.8), Inches(4.5), Inches(1.2)]

hdr_cells2 = table_tests.rows[0].cells
for i, title in enumerate(tc_headers):
    hdr_cells2[i].text = title
    set_cell_background(hdr_cells2[i], "1F4E79")
    set_cell_margins(hdr_cells2[i], top=100, bottom=100, left=100, right=100)
    p = hdr_cells2[i].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in p.runs:
        run.font.name = "Arial"
        run.font.size = Pt(10)
        run.font.bold = True
        run.font.color.rgb = RGBColor(255, 255, 255)

tests_data = [
    ("T1", "giỏ mới: tongTien() là 0, layDanhSach() là mảng rỗng", "✅ Đạt"),
    ("T2", "them trả về true, dòng mới có đủ id, ten, loai, gia, soLuong", "✅ Đạt"),
    ("T3", "thêm cùng một sản phẩm hai lần → vẫn 1 dòng, số lượng cộng dồn (100 + 50 = 150)", "✅ Đạt"),
    ("T4", "số lượng 0, âm, lẻ (2.5), NaN, null, undefined → false, giỏ không đổi", "✅ Đạt"),
    ("T5", "số lượng là CHUỖI \"50\" → false; số lượng cũ vẫn là number 100", "✅ Đạt"),
    ("T6", "DANH_MUC không bị thay đổi sau khi thêm vào giỏ", "✅ Đạt"),
    ("T7", "giỏ không giữ tham chiếu tới object sản phẩm bên ngoài", "✅ Đạt"),
    ("T8", "xoa: có → true và mất dòng; không có → false và giữ nguyên", "✅ Đạt"),
    ("T9", "tongTien của giỏ mẫu là 340000", "✅ Đạt"),
    ("T10", "thongKeTheoLoai: { \"thiep cuoi\": 240000, \"phu kien\": 100000 }, không thừa khoá", "✅ Đạt"),
    ("T11", "thongKeTheoLoai của giỏ rỗng là object rỗng", "✅ Đạt"),
    ("T12", "dongDatNhat: dòng A, thanhTien 150000; giỏ rỗng → null", "✅ Đạt"),
    ("T13", "dongDatNhat: hai dòng bằng nhau → lấy dòng ĐỨNG TRƯỚC", "✅ Đạt"),
    ("T14", "layDanhSach có thanhTien cho từng dòng", "✅ Đạt"),
    ("T15", "sửa kết quả của layDanhSach / dongDatNhat KHÔNG làm đổi giỏ", "✅ Đạt"),
    ("T16", "xuatJSON trả về chuỗi; parse ra mảng, mỗi phần tử có ĐÚNG 5 khoá id, ten, loai, gia, soLuong", "✅ Đạt"),
    ("T17", "napGioHang trả về giỏ dùng được phương thức: tongTien() là 340000", "✅ Đạt"),
    ("T18", "giỏ nạp lại vẫn cộng dồn đúng và độc lập với giỏ gốc", "✅ Đạt"),
    ("T19", "napGioHang(\"[]\") → giỏ rỗng, tongTien() là 0", "✅ Đạt"),
    ("T20", "xuất → nạp → xuất cho ra đúng chuỗi ban đầu", "✅ Đạt")
]

for row_idx, data in enumerate(tests_data):
    row_cells = table_tests.add_row().cells
    bg_color = "F2F2F2" if row_idx % 2 == 1 else "FFFFFF"
    for col_idx, text in enumerate(data):
        row_cells[col_idx].text = text
        set_cell_background(row_cells[col_idx], bg_color)
        set_cell_margins(row_cells[col_idx], top=70, bottom=70, left=90, right=90)
        p = row_cells[col_idx].paragraphs[0]
        if col_idx in [0, 2]:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in p.runs:
            run.font.name = "Arial"
            run.font.size = Pt(9.5)
            if col_idx == 2:
                run.font.bold = True
                run.font.color.rgb = RGBColor(0, 128, 0)

for row in table_tests.rows:
    for i, w in enumerate(tc_col_widths):
        row.cells[i].width = w

doc.add_paragraph()
p_summary = doc.add_paragraph()
r_sum = p_summary.add_run("KẾT LUẬN: BÀI B ĐẠT 20/20 TESTCASES (100% PASS).")
r_sum.font.name = "Arial"
r_sum.font.size = Pt(11)
r_sum.font.bold = True
r_sum.font.color.rgb = RGBColor(0, 128, 0)

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
r_idx_h = p_script_idx.add_run("• Trang news.html (Bài 2A):\n")
r_idx_h.font.bold = True
p_script_idx.add_run("   <script src=\"js/ho-tro.js\"></script>\n"
                       "   <script src=\"js/bai-a-tin-tuc.js\"></script>\n"
                       "-> Phân tích:\n"
                       "   1. ho-tro.js được nạp trước để định nghĩa các hàm veDanhSachTin(dsBai) và hienThi(noiDung) vào phạm vi toàn cục.\n"
                       "   2. bai-a-tin-tuc.js được nạp sau sẽ thực thi việc phân tích chuỗi JSON, xử lý mảng bài viết và gọi veDanhSachTin(), hienThi() để cập nhật giao diện. Nếu nạp ngược lại, trình duyệt sẽ quăng lỗi ReferenceError: veDanhSachTin is not defined.\n\n")

r_ab_h = p_script_idx.add_run("• Trang index.html (Bài 2B):\n")
r_ab_h.font.bold = True
p_script_idx.add_run("   <script src=\"js/ho-tro.js\"></script>\n"
                      "   <script src=\"js/bai-b-gio-hang.js\"></script>\n"
                      "   <script src=\"js/kiem-tra.js\"></script>\n"
                      "-> Phân tích:\n"
                      "   1. ho-tro.js được nạp đầu tiên để chuẩn bị sẵn hàm vẽ giao diện veGioHang() và hienThi().\n"
                      "   2. bai-b-gio-hang.js nạp thứ hai nhằm khai báo hằng số DANH_MUC, định nghĩa class GioHang và hàm napGioHang(), sau đó chạy đoạn mã demo thêm sản phẩm vào giỏ và gọi veGioHang().\n"
                      "   3. kiem-tra.js được nạp sau cùng để thực thi bộ 20 testcases tự động nhằm kiểm thử toàn diện class GioHang và hàm napGioHang() đã được định nghĩa từ bước 2. Nếu kiem-tra.js nạp trước, các lời gọi new GioHang() trong testcases sẽ bị lỗi ReferenceError: GioHang is not defined.")

# Save document
output_path = os.path.join(current_dir, "Bao_cao_Bai_tap_2_JS.docx")
doc.save(output_path)
shutil.copyfile(output_path, os.path.join(current_dir, "Bao_cao_Bai_tap_2_JS.doc"))
print(f"File created successfully at: {output_path}")
