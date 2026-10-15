/* BÀI B — VIẾT CODE: Lọc tin và đăng tin nhanh (news.html)
   Viết code theo đặc tả.
   Không sửa news.html, không sửa css; chỉ viết JavaScript trong file này. */

// ---------- ĐÃ CHO SẴN (không sửa) ----------
// Chặn việc gửi form tìm kiếm khi nhấn Enter, để trang không bị tải lại.
document.querySelector("#searchbar form").addEventListener("submit", (e) => e.preventDefault());

// ---------- PHẦN BẠN VIẾT ----------

const oTimKiem = document.getElementById("o-tim-kiem");
const list = document.getElementById("list");
const boDem = document.getElementById("bo-dem");
const khongThay = document.getElementById("khong-thay");
const oTieuDe = document.getElementById("tin-tieu-de");
const oNoiDung = document.getElementById("tin-noi-dung");
const nutDang = document.getElementById("nut-dang");
const baoLoi = document.getElementById("bao-loi");

// Hàm cập nhật bộ lọc và bộ đếm
function capNhatBoLoc() {
	const tuKhoa = oTimKiem.value;
	const dsMuc = Array.from(list.querySelectorAll("li"));
	let soHien = 0;

	dsMuc.forEach(muc => {
		const h3 = muc.querySelector("h3");
		const p = muc.querySelector("p");
		const vanBan = (h3 ? h3.textContent : "") + " " + (p ? p.textContent : "");
		const khop = tuKhoa === "" || vanBan.toLowerCase().includes(tuKhoa.toLowerCase());

		if (khop) {
			muc.classList.remove("an");
			soHien++;
		} else {
			muc.classList.add("an");
		}
	});

	if (soHien === 0) {
		khongThay.classList.remove("an");
		khongThay.textContent = `Không tìm thấy tin nào cho "${tuKhoa}"`;
	} else {
		khongThay.classList.add("an");
	}

	boDem.textContent = `Đang hiển thị ${soHien}/${dsMuc.length} tin`;
}

// F1 + F2 + F3. Bộ đếm, lọc tin khi gõ và thông báo không tìm thấy
oTimKiem.addEventListener("keyup", capNhatBoLoc);

// Cập nhật bộ đếm ngay khi tải trang
capNhatBoLoc();

// F4. Nút "Read more >>" / "Thu gọn <<" — MỘT trình lắng nghe đặt trên #list
list.addEventListener("click", (e) => {
	if (!e.target.classList.contains("more") || e.target.tagName !== "BUTTON") {
		return;
	}
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
});

// F5. Nút #nut-dang: kiểm tra dữ liệu, tạo <li> mới, chèn vào đầu #list
nutDang.addEventListener("click", () => {
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
});
