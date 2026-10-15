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
title_run = title_p.add_run("BÁO CÁO BÀI TẬP 3: JAVASCRIPT")
title_run.font.name = "Arial"
title_run.font.size = Pt(20)
title_run.font.bold = True
title_run.font.color.rgb = RGBColor(31, 78, 121)

sub_p = doc.add_paragraph()
sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
sub_run = sub_p.add_run("Thư Viện Ảnh Tương Tác & Lọc - Đăng Tin Nhanh (Blakletterpress)")
sub_run.font.name = "Arial"
sub_run.font.size = Pt(13)
sub_run.font.italic = True
sub_run.font.color.rgb = RGBColor(89, 89, 89)

doc.add_paragraph() # Spacer

# ==================== PHẦN 1: BÀI 3A ====================
h1 = doc.add_paragraph()
r = h1.add_run("1. BÀI 3A — ĐIỀN VÀO CHỖ TRỐNG THƯ VIỆN ẢNH (js/bai-a-gallery.js)")
r.font.name = "Arial"
r.font.size = Pt(14)
r.font.bold = True
r.font.color.rgb = RGBColor(31, 78, 121)

p_intro_3a = doc.add_paragraph()
p_intro_3a.add_run("Dưới đây là bảng tổng hợp chi tiết đáp án cho 27 chỗ trống trong file ")
r_code = p_intro_3a.add_run("js/bai-a-gallery.js")
r_code.font.bold = True
p_intro_3a.add_run(", bao gồm các phương thức DOM, thuộc tính, từ khóa, sự kiện, kết quả dự đoán và giải thích cơ chế:")

# Table for 27 blanks
table_3a = doc.add_table(rows=1, cols=4)
table_3a.alignment = WD_TABLE_ALIGNMENT.CENTER
table_3a.autofit = False

headers = ["STT", "Đáp án điền", "Loại", "Giải thích chi tiết & cơ chế JS / DOM"]
col_widths = [Inches(0.6), Inches(1.8), Inches(0.9), Inches(3.2)]

hdr_cells = table_3a.rows[0].cells
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

data_3a = [
    ("(1)", "querySelector", "Code", "document.querySelector(\"#gallery > div > img\") tìm phần tử đầu tiên trong DOM khớp với bộ chọn CSS."),
    ("(2)", "querySelectorAll", "Code", "document.querySelectorAll(\"#gallery ul li img\") tìm tất cả các ảnh nhỏ khớp bộ chọn và trả về một NodeList."),
    ("(3)", "getElementById", "Code", "document.getElementById(\"o-tim-kiem\") tìm phần tử theo thuộc tính id duy nhất."),
    ("(4)", "6", "Dự đoán", "Trong #gallery ul có 6 thẻ <li>, mỗi thẻ chứa 1 ảnh nhỏ <img>, do đó dsAnhNho.length = 6."),
    ("(5)", "3", "Dự đoán", "KhungAnhNho.firstChild là khoảng trắng (xuống dòng và thụt lề) trước thẻ <li> đầu tiên; đây là TextNode có nodeType = 3 (Node.TEXT_NODE)."),
    ("(6)", "13", "Dự đoán", "khungAnhNho chứa 6 Element nodes (các thẻ <li>) và 7 Text nodes (khoảng trắng trước, xen kẽ và sau các <li>), tổng cộng 13 childNodes."),
    ("(7)", "parentNode", "Code", "Đi từ phần tử <img> lên phần tử cha chứa nó là thẻ <li> (có thuộc tính nodeName là 'LI')."),
    ("(8)", "createElement", "Code", "document.createElement(\"p\") khởi tạo một phần tử HTML <p> mới trong bộ nhớ cây DOM."),
    ("(9)", "textContent", "Code", "Gán văn bản thuần dạng chuỗi, bảo đảm trình duyệt không diễn giải thành các thẻ HTML."),
    ("(10)", "appendChild", "Code", "Chèn phần tử chuThich vào cuối danh sách con của thẻ <div> đang bao bọc ảnh lớn."),
    ("(11)", "!==", "Code", "Toán tử so sánh khác nghiêm ngặt (!==); nếu phần tử được bấm không phải thẻ IMG thì dừng lại."),
    ("(12)", "return", "Code", "Lệnh dừng thực thi hàm doiAnh khi người dùng bấm vào khoảng trống của <ul> hoặc mép <li>, tránh phát sinh lỗi."),
    ("(13)", "target", "Code", "e.target đại diện cho phần tử phát sinh sự kiện click (chính là thẻ <img> vừa được bấm)."),
    ("(14)", "of", "Code", "Cú pháp vòng lặp for (const anh of dsAnhNho) duyệt qua từng phần tử trong NodeList dsAnhNho."),
    ("(15)", "className", "Code", "Gán lớp \"dang-chon\" cho thẻ <li> cha của ảnh vừa bấm để thêm viền highlight, tương ứng với đoạn xóa lớp ở dòng 46."),
    ("(16)", "addEventListener", "Code", "Đăng ký lắng nghe sự kiện \"click\" trên thẻ cha khungAnhNho theo cơ chế Event Delegation."),
    ("(17)", "doiAnh", "Code", "Truyền tên hàm (tham chiếu callback) doiAnh để trình duyệt gọi khi sự kiện click xảy ra."),
    ("(18)", "(B) báo TypeError ngay khi tải trang", "Dự đoán", "Viết doiAnh() với dấu ngoặc đơn sẽ gọi hàm ngay lúc đăng ký; khi đó e là undefined, truy cập e.target.nodeName sẽ gây lỗi TypeError."),
    ("(19)", "focus", "Code", "Sự kiện \"focus\" kích hoạt khi người dùng nhấp chuột hoặc chuyển tiêu điểm vào ô tìm kiếm."),
    ("(20)", "removeEventListener", "Code", "Gỡ bỏ trình lắng nghe sự kiện click đã đăng ký trên khungAnhNho khi bấm nút 'Khoá thư viện ảnh'."),
    ("(21)", "doiAnh", "Code", "Truyền đúng tham chiếu hàm callback doiAnh ban đầu để gỡ bỏ chính xác."),
    ("(22)", "click", "Dự đoán", "Thuộc tính e.type lưu tên loại sự kiện đang được kích hoạt, ở đây là 'click'."),
    ("(23)", "nut-khoa", "Dự đoán", "Thuộc tính e.target.id lưu giá trị id của nút người dùng vừa bấm, là 'nut-khoa'."),
    ("(24)", "không", "Dự đoán", "Hàm mũi tên ẩn danh (e) => doiAnh(e) tạo ra một instance hàm hoàn toàn mới trong ô nhớ; removeEventListener không thể xác định và không gỡ được."),
    ("(25)", "prepend", "Code", "Phương thức prepend() chèn một hoặc nhiều node mới vào ĐẦU danh sách phần tử con của dsBaiViet."),
    ("(26)", "hiện NGUYÊN cả cặp thẻ <b>…</b>", "Dự đoán", "Vì tieuDe được gán bằng textContent nên toàn bộ chuỗi được coi là ký tự văn bản thuần, không diễn giải in đậm thẻ <b>."),
    ("(27)", "3", "Dự đoán", "Danh sách #posts ul ban đầu có 2 thẻ <li> mẫu, sau khi prepend thêm 1 thẻ thì tổng số thẻ <li> là 3.")
]

for row_idx, data in enumerate(data_3a):
    row_cells = table_3a.add_row().cells
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

for row in table_3a.rows:
    for i, w in enumerate(col_widths):
        row.cells[i].width = w

doc.add_paragraph() # Spacer

# Hiệu ứng kiểm tra
h2_eff = doc.add_paragraph()
r = h2_eff.add_run("Các hiệu ứng hoạt động trên trang index.html:")
r.font.name = "Arial"
r.font.size = Pt(11)
r.font.bold = True

eff_text = """1. Dưới ảnh lớn có dòng chú thích: 'Bấm vào ảnh nhỏ để đổi ảnh lớn.'
2. Bấm một ảnh nhỏ: ảnh lớn đổi src, chú thích cập nhật 'Đang xem: meeting.jpg' (hoặc team-large.jpg), ảnh nhỏ được chọn có class 'dang-chon' (hiện viền).
3. Bấm khoảng trống giữa các ảnh nhỏ: không có gì xảy ra, console không có lỗi.
4. Bấm vào ô tìm kiếm: chữ gợi ý 'Enter search query' biến mất; rời ô khi chưa nhập: chữ gợi ý tự phục hồi.
5. Bấm 'Khoá thư viện ảnh': gỡ listener, bấm ảnh nhỏ không còn đổi ảnh lớn, chú thích đổi 'Thư viện ảnh đã khoá.'
6. Cột bên phải (#posts): có thêm 1 mục ở đầu danh sách với nội dung hiển thị nguyên văn 'Khuyến mãi <b>thiệp cưới</b>'."""

p_eff = doc.add_paragraph()
r_ef = p_eff.add_run(eff_text)
r_ef.font.name = "Consolas"
r_ef.font.size = Pt(9.5)
r_ef.font.bold = True
r_ef.font.color.rgb = RGBColor(0, 102, 0)

# Mã nguồn hoàn chỉnh bài 3a
h2_3a = doc.add_paragraph()
r = h2_3a.add_run("Mã nguồn hoàn chỉnh js/bai-a-gallery.js:")
r.font.name = "Arial"
r.font.size = Pt(11)
r.font.bold = True

current_dir = os.path.dirname(os.path.abspath(__file__))
with open(os.path.join(current_dir, "js", "bai-a-gallery.js"), "r", encoding="utf-8") as f:
    code_content_3a = f.read()

code_p3a = doc.add_paragraph()
r_code_3a = code_p3a.add_run(code_content_3a)
r_code_3a.font.name = "Consolas"
r_code_3a.font.size = Pt(8.5)
r_code_3a.font.color.rgb = RGBColor(40, 40, 40)

doc.add_page_break()

# ==================== PHẦN 2: BÀI 3B ====================
h1_b = doc.add_paragraph()
r = h1_b.add_run("2. BÀI 3B — LỌC TIN VÀ ĐĂNG TIN NHANH (js/bai-b-tin-nhanh.js)")
r.font.name = "Arial"
r.font.size = Pt(14)
r.font.bold = True
r.font.color.rgb = RGBColor(31, 78, 121)

p_intro_3b = doc.add_paragraph()
p_intro_3b.add_run("Bài 3b yêu cầu cài đặt đầy đủ 5 chức năng (F1 đến F5) trong file ")
r_code = p_intro_3b.add_run("js/bai-b-tin-nhanh.js")
r_code.font.bold = True
p_intro_3b.add_run(" để vượt qua toàn bộ 18 testcases tự động (T1 đến T18) của ")
r_kt = p_intro_3b.add_run("js/kiem-tra.js")
r_kt.font.bold = True
p_intro_3b.add_run(" trên trang ")
r_pg = p_intro_3b.add_run("news.html")
r_pg.font.bold = True
p_intro_3b.add_run(":")

# 5 features explanation
features = [
    ("F1. Bộ đếm tin tức (#bo-dem)",
     "• Mục đích: Luôn hiển thị trạng thái 'Đang hiển thị x/y tin', trong đó x là số tin đang hiển thị (không chứa lớp 'an') và y là tổng số tin trong #list.\n"
     "• Thực thi: Hàm capNhatBoLoc() tự động cập nhật bộ đếm ngay khi trang tải xong, sau mỗi lần lọc gõ phím và sau mỗi lần thêm tin mới thành công.",
     """boDem.textContent = `Đang hiển thị ${soHien}/${dsMuc.length} tin`;"""),

    ("F2. Lọc tin khi gõ (#o-tim-kiem)",
     "• Sự kiện: Lắng nghe sự kiện 'keyup' trên ô tìm kiếm #o-tim-kiem.\n"
     "• Cơ chế so khớp: Lấy nội dung kết hợp của tiêu đề <h3> và đoạn văn <p>, chuyển về chữ thường bằng toLowerCase() và so sánh chuỗi với tuKhoa.toLowerCase() thông qua includes().\n"
     "• Trạng thái: Tin không khớp nhận thêm lớp 'an'; tin khớp được bỏ lớp 'an'. Khi từ khoá rỗng, toàn bộ tin đều được hiển thị.",
     """dsMuc.forEach(muc => {
	const h3 = muc.querySelector("h3");
	const p = muc.querySelector("p");
	const vanBan = (h3 ? h3.textContent : "") + " " + (p ? p.textContent : "");
	const khop = tuKhoa === "" || vanBan.toLowerCase().includes(tuKhoa.toLowerCase());
	if (khop) muc.classList.remove("an");
	else muc.classList.add("an");
});"""),

    ("F3. Thông báo không tìm thấy (#khong-thay)",
     "• Kiểm tra kết quả lọc: Nếu soHien === 0 (không có tin nào khớp), thẻ #khong-thay được gỡ bỏ lớp 'an' và hiển thị nội dung: Không tìm thấy tin nào cho \"<từ khoá>\".\n"
     "• Bảo mật XSS & Nguyên văn: Sử dụng thuộc tính textContent để gán nội dung thông báo, đảm bảo các thẻ HTML nhập vào như <b>x</b> được hiển thị nguyên bản dạng chuỗi, không bị parse thành thẻ HTML.",
     """if (soHien === 0) {
	khongThay.classList.remove("an");
	khongThay.textContent = `Không tìm thấy tin nào cho "${tuKhoa}"`;
} else {
	khongThay.classList.add("an");
}"""),

    ("F4. Mở rộng / Thu gọn nội dung (Event Delegation)",
     "• Kiến trúc lắng nghe: Chỉ đăng ký duy nhất MỘT trình lắng nghe sự kiện 'click' trên thẻ cha #list.\n"
     "• Kiểm tra đối tượng: Chỉ xử lý khi e.target là phần tử <button> có class 'more'. Bấm vào tiêu đề, nội dung hay khoảng trống khác không gây ảnh hưởng.\n"
     "• Chuyển đổi trạng thái: Bấm lần đầu, thẻ <p> nhận lớp 'day-du' và chữ nút chuyển thành 'Thu gọn <<'. Bấm lần tiếp theo, gỡ lớp 'day-du' và nhãn nút trở lại 'Read more >>'. Các tin mới đăng tự động hưởng cơ chế này.",
     """list.addEventListener("click", (e) => {
	if (!e.target.classList.contains("more") || e.target.tagName !== "BUTTON") return;
	const muc = e.target.closest("li");
	if (!muc) return;
	const doan = muc.querySelector("p");
	if (!doan) return;
	if (doan.classList.contains("day-du")) {
		doan.classList.remove("day-du");
		e.target.textContent = "Read more >>";
	} else {
		doan.classList.add("day-du");
		e.target.textContent = "Thu gọn <<";
	}
});"""),

    ("F5. Đăng tin nhanh (#nut-dang)",
     "• Kiểm tra hợp lệ: Kiểm tra tiêu đề và nội dung; nếu rỗng (sau trim) thì hiển thị cảnh báo 'Vui lòng nhập đủ tiêu đề và nội dung.' vào #bao-loi và dừng lại.\n"
     "• Tạo thẻ DOM: Tạo <li> mới gồm đầy đủ 4 phần tử con: <span class=\"time\"> (lấy ngày hiện tại qua hàm homNay()), <h3>, <p>, và <button type=\"button\" class=\"more\">Read more >></button>.\n"
     "• Chèn vào đầu: Dùng list.prepend(mucMoi) để đưa tin mới lên đầu danh sách.\n"
     "• Dọn dẹp & Tái áp dụng bộ lọc: Xóa trắng hai ô nhập liệu, xóa #bao-loi, gọi lại hàm capNhatBoLoc() để tin mới tự động tuân theo bộ lọc hiện thời và cập nhật lại bộ đếm.",
     """nutDang.addEventListener("click", () => {
	const tieuDe = oTieuDe.value;
	const noiDung = oNoiDung.value;
	if (tieuDe.trim() === "" || noiDung.trim() === "") {
		baoLoi.textContent = "Vui lòng nhập đủ tiêu đề và nội dung.";
		return;
	}
	baoLoi.textContent = "";
	const mucMoi = document.createElement("li");
	const timeSpan = document.createElement("span");
	timeSpan.className = "time";
	timeSpan.textContent = homNay();
	const h3 = document.createElement("h3");
	h3.textContent = tieuDe;
	const p = document.createElement("p");
	p.textContent = noiDung;
	const nutMore = document.createElement("button");
	nutMore.type = "button";
	nutMore.className = "more";
	nutMore.textContent = "Read more >>";
	mucMoi.append(timeSpan, h3, p, nutMore);
	list.prepend(mucMoi);
	oTieuDe.value = "";
	oNoiDung.value = "";
	capNhatBoLoc();
});""")
]

for title, desc, code in features:
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

# Bảng kết quả 18 testcases
doc.add_paragraph()
h2_tc = doc.add_paragraph()
r = h2_tc.add_run("Bảng kết quả kiểm tra tự động 18 testcases (js/kiem-tra.js):")
r.font.name = "Arial"
r.font.size = Pt(11)
r.font.bold = True

table_tests3 = doc.add_table(rows=1, cols=3)
table_tests3.alignment = WD_TABLE_ALIGNMENT.CENTER
table_tests3.autofit = False

tc_headers3 = ["STT", "Nội dung kiểm tra", "Kết quả"]
tc_col_widths3 = [Inches(0.8), Inches(4.5), Inches(1.2)]

hdr_cells3 = table_tests3.rows[0].cells
for i, title in enumerate(tc_headers3):
    hdr_cells3[i].text = title
    set_cell_background(hdr_cells3[i], "1F4E79")
    set_cell_margins(hdr_cells3[i], top=100, bottom=100, left=100, right=100)
    p = hdr_cells3[i].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in p.runs:
        run.font.name = "Arial"
        run.font.size = Pt(10)
        run.font.bold = True
        run.font.color.rgb = RGBColor(255, 255, 255)

tests_data3 = [
    ("T1", "không có từ khoá: hiện tất cả, bộ đếm \"Đang hiển thị N/N tin\"", "✅ Đạt"),
    ("T2", "gõ \"cotton\": chỉ còn các tin có chữ cotton trong tiêu đề hoặc nội dung", "✅ Đạt"),
    ("T3", "không phân biệt hoa thường: \"CoTToN\" cho kết quả như \"cotton\"", "✅ Đạt"),
    ("T4", "từ khoá chỉ có trong nội dung (\"phong bì\") vẫn tìm thấy", "✅ Đạt"),
    ("T5", "bộ đếm cập nhật theo kết quả lọc", "✅ Đạt"),
    ("T6", "không tin nào khớp: #khong-thay hiện ra với đúng nội dung, bộ đếm 0/N", "✅ Đạt"),
    ("T7", "từ khoá chứa thẻ HTML được hiện nguyên văn, không thành phần tử", "✅ Đạt"),
    ("T8", "xoá từ khoá: hiện lại tất cả, #khong-thay ẩn, cả trang chỉ có một #khong-thay", "✅ Đạt"),
    ("T9", "bấm \"Read more >>\": đoạn văn có lớp day-du, nút đổi thành \"Thu gọn <<\"; bấm lần nữa thì trở lại", "✅ Đạt"),
    ("T10", "bấm vào tiêu đề hoặc đoạn văn (không phải nút) thì không có gì thay đổi", "✅ Đạt"),
    ("T11", "mở tin thứ nhất không làm tin thứ hai mở theo", "✅ Đạt"),
    ("T12", "thiếu tiêu đề hoặc thiếu nội dung: báo lỗi đúng câu, không thêm tin", "✅ Đạt"),
    ("T13", "đăng hợp lệ: <li> mới nằm ĐẦU danh sách, đủ span.time, h3, p, button.more", "✅ Đạt"),
    ("T14", "sau khi đăng: hai ô nhập trống, hết báo lỗi, bộ đếm tăng", "✅ Đạt"),
    ("T15", "nội dung người dùng nhập được hiện nguyên văn, không thành thẻ HTML", "✅ Đạt"),
    ("T16", "nút \"Read more >>\" của tin VỪA ĐĂNG cũng hoạt động", "✅ Đạt"),
    ("T17", "đang lọc \"cotton\" mà đăng tin không chứa cotton: tin mới bị ẩn, bộ đếm vẫn đúng", "✅ Đạt"),
    ("T18", "đang lọc \"cotton\" mà đăng tin CÓ chữ Cotton: tin mới hiện ra", "✅ Đạt")
]

for row_idx, data in enumerate(tests_data3):
    row_cells = table_tests3.add_row().cells
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

for row in table_tests3.rows:
    for i, w in enumerate(tc_col_widths3):
        row.cells[i].width = w

doc.add_paragraph()
p_summary3 = doc.add_paragraph()
r_sum3 = p_summary3.add_run("KẾT LUẬN: BÀI B ĐẠT 18/18 TESTCASES (100% PASS).")
r_sum3.font.name = "Arial"
r_sum3.font.size = Pt(11)
r_sum3.font.bold = True
r_sum3.font.color.rgb = RGBColor(0, 128, 0)

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
r_idx_h = p_script_idx.add_run("• Trang index.html (Bài 3A):\n")
r_idx_h.font.bold = True
p_script_idx.add_run("   <script src=\"js/bai-a-gallery.js\"></script>\n"
                       "-> Phân tích: Script được đặt ở cuối body sau khi toàn bộ cấu trúc cây DOM (#gallery, div, img, ul, li, #posts) đã được dựng hoàn chỉnh. Khi script thực thi, các lệnh querySelector() và getElementById() sẽ tìm thấy ngay các phần tử mà không bị trả về null.\n\n")

r_ab_h = p_script_idx.add_run("• Trang news.html (Bài 3B):\n")
r_ab_h.font.bold = True
p_script_idx.add_run("   <script src=\"js/ho-tro.js\"></script>\n"
                      "   <script src=\"js/bai-b-tin-nhanh.js\"></script>\n"
                      "   <script src=\"js/kiem-tra.js\"></script>\n"
                      "-> Phân tích:\n"
                      "   1. ho-tro.js được nạp đầu tiên để đăng ký hàm homNay() trả về chuỗi ngày tháng định dạng dd-mm-yyyy vào phạm vi toàn cục.\n"
                      "   2. bai-b-tin-nhanh.js nạp thứ hai nhằm khởi tạo các trình lắng nghe sự kiện keyup cho ô tìm kiếm, click cho bộ ủy quyền sự kiện trên #list và nút đăng tin #nut-dang, đồng thời gọi homNay() khi tạo bài mới. Nếu nạp trước ho-tro.js thì khi đăng tin sẽ bị lỗi 'homNay is not defined'.\n"
                      "   3. kiem-tra.js được nạp sau cùng để gắn sự kiện click vào nút #nut-kiem-tra, chuẩn bị sẵn sàng chạy 18 testcases giả lập thao tác người dùng.")

# Save document
output_path = os.path.join(current_dir, "Bao_cao_Bai_tap_3_JS.docx")
doc.save(output_path)
shutil.copyfile(output_path, os.path.join(current_dir, "Bao_cao_Bai_tap_3_JS.doc"))
print(f"File created successfully at: {output_path}")
