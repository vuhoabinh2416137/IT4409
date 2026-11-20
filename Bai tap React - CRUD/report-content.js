(function () {
    "use strict";

    function renderReactCrudReport(containerId, options) {
        var container = document.getElementById(containerId);
        if (!container) return;

        var opts = options || {};
        var isStandalone = Boolean(opts.standalone);

        var actionButton = isStandalone
            ? '<a class="btn-outline-link" href="index.html">&#8592; Mở Ứng dụng React CRUD</a>'
            : '<a class="btn-outline-link" href="bao-cao.html">Mở trang Báo cáo độc lập (bao-cao.html)</a>';

        container.innerHTML = `
            <div class="report-header">
                <h2 class="report-title">BÁO CÁO KẾT QUẢ BÀI TẬP: REACT CƠ BẢN (CRUD)</h2>
                <p class="report-subtitle">
                    Ứng dụng Quản lý Người dùng (Create - Read - Update - Delete) với React 18, Hooks &amp; State Lifting
                </p>
                <div class="report-meta">
                    <span class="meta-pill">Học phần: IT4409</span>
                    <span class="meta-pill">Sinh viên: Vũ Hòa Bình (MSSV: 202416137)</span>
                    <span class="meta-pill">Host: Vercel</span>
                </div>
                <div class="view-actions" style="justify-content: center; margin-top: 16px;">
                    ${actionButton}
                    <a class="btn-download" href="Bao_cao_Bai_tap_React_CRUD.docx" download>
                        &#11015; Tải Báo cáo Word (.docx)
                    </a>
                    <a class="btn-outline-link" href="Bao_cao_Bai_tap_React_CRUD.doc" download>
                        &#11015; Tải bản (.doc)
                    </a>
                </div>
            </div>

            <!-- PHẦN 1: SƠ ĐỒ COMPONENT -->
            <section class="report-section" id="sec-component-tree">
                <h2 class="report-h2">1. Sơ đồ cây Component (Component Tree)</h2>
                <p class="report-text">
                    Ứng dụng Quản lý người dùng được tổ chức theo mô hình Functional Component của React 18, gồm 1 component gốc
                    (<code class="code-inline">App</code>) và 3 component con ngang hàng
                    (<code class="code-inline">SearchForm</code>, <code class="code-inline">AddUser</code>, <code class="code-inline">ResultTable</code>)
                    được gắn vào <code class="code-inline">&lt;div id="root"&gt;</code> qua <code class="code-inline">ReactDOM.createRoot()</code>.
                </p>

                <div class="diagram-card">
                    <img src="component_tree.png" alt="Sơ đồ cây Component (Component Tree) của ứng dụng React CRUD">
                    <div class="diagram-caption">Hình 1.1: Sơ đồ cây Component (Component Tree) — Phân cấp cha/con, Props và State.</div>
                </div>

                <pre class="code-block">index.html (&lt;div id="root"&gt;&lt;/div&gt;)
└── &lt;App /&gt;  [State dùng chung: kw, newUser]
    ├── &lt;SearchForm onChangeValue={setKeyword} /&gt;        [Stateless - nhận input tìm kiếm]
    ├── &lt;AddUser onAdd={setNewUser} /&gt;                   [State nội bộ: adding, user]
    └── &lt;ResultTable keyword={kw} user={newUser} onAdded={() =&gt; setNewUser(null)} /&gt;
                                                         [State nội bộ: users, loading, error, editing]</pre>

                <div class="report-table-wrap">
                    <table>
                        <thead>
                            <tr>
                                <th>Component</th>
                                <th>Cấp / Vai trò</th>
                                <th>State nội bộ (Local State)</th>
                                <th>Nhiệm vụ chính</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td><strong>App</strong></td>
                                <td>Root (Cha)</td>
                                <td><code class="code-inline">kw</code> (string), <code class="code-inline">newUser</code> (object | null)</td>
                                <td>Lưu trữ toàn bộ state dùng chung và truyền props (dữ liệu &amp; callback) xuống các component con.</td>
                            </tr>
                            <tr>
                                <td><strong>SearchForm</strong></td>
                                <td>Child (Con)</td>
                                <td>Không có (Stateless)</td>
                                <td>Hiển thị ô <code class="code-inline">&lt;input&gt;</code> tìm kiếm; gọi <code class="code-inline">onChangeValue(e.target.value)</code> khi người dùng gõ phím.</td>
                            </tr>
                            <tr>
                                <td><strong>AddUser</strong></td>
                                <td>Child (Con)</td>
                                <td><code class="code-inline">adding</code> (boolean), <code class="code-inline">user</code> (object lồng <code class="code-inline">address</code>)</td>
                                <td>Quản lý Modal thêm người dùng mới, cập nhật object lồng nhau bằng Spread Operator (<code class="code-inline">...</code>) và gửi lên <code class="code-inline">App</code>.</td>
                            </tr>
                            <tr>
                                <td><strong>ResultTable</strong></td>
                                <td>Child (Con)</td>
                                <td><code class="code-inline">users</code> (array), <code class="code-inline">loading</code>, <code class="code-inline">error</code>, <code class="code-inline">editing</code></td>
                                <td>Tải danh sách từ API bằng <code class="code-inline">useEffect</code>; lắng nghe prop <code class="code-inline">user</code>; lọc theo <code class="code-inline">keyword</code>; Sửa (Deep Copy) và Xóa.</td>
                            </tr>
                        </tbody>
                    </table>
                </div>
            </section>

            <!-- PHẦN 2: SƠ ĐỒ LUỒNG DỮ LIỆU & STATE LIFTING -->
            <section class="report-section" id="sec-state-lifting">
                <h2 class="report-h2">2. Sơ đồ luồng dữ liệu (State Lifting) &amp; Giải thích đặt State tại App</h2>
                <p class="report-text">
                    Ứng dụng áp dụng mô hình <strong>Luồng dữ liệu một chiều (One-Way Data Flow)</strong> kết hợp kỹ thuật
                    <strong>Nâng State lên cha chung (State Lifting)</strong> để đồng bộ dữ liệu giữa các component ngang hàng.
                </p>

                <div class="diagram-card">
                    <img src="state_lifting_flow.png" alt="Sơ đồ luồng dữ liệu State Lifting và One-Way Data Flow">
                    <div class="diagram-caption">Hình 2.1: Sơ đồ luồng dữ liệu hai nhánh — Tìm kiếm (SearchForm &rarr; App &rarr; ResultTable) và Thêm mới (AddUser &rarr; App &rarr; ResultTable &rarr; App).</div>
                </div>

                <h3 class="report-h3">2.1. Hai luồng dữ liệu chính</h3>
                <ul class="report-list">
                    <li>
                        <strong>Luồng Tìm kiếm:</strong> <code class="code-inline">SearchForm</code> (sự kiện <code class="code-inline">onChange</code>)
                        &rarr; gọi callback <code class="code-inline">onChangeValue(text)</code> &rarr; cập nhật <code class="code-inline">kw</code> trên <code class="code-inline">App</code>
                        &rarr; <code class="code-inline">App</code> truyền prop <code class="code-inline">keyword={kw}</code> xuống <code class="code-inline">ResultTable</code>
                        &rarr; lọc danh sách hiển thị theo <code class="code-inline">name</code> và <code class="code-inline">username</code>.
                    </li>
                    <li>
                        <strong>Luồng Thêm người dùng:</strong> <code class="code-inline">AddUser</code> (bấm nút Thêm)
                        &rarr; gọi callback <code class="code-inline">onAdd(user)</code> &rarr; cập nhật <code class="code-inline">newUser</code> trên <code class="code-inline">App</code>
                        &rarr; <code class="code-inline">App</code> truyền prop <code class="code-inline">user={newUser}</code> xuống <code class="code-inline">ResultTable</code>
                        &rarr; <code class="code-inline">useEffect([user])</code> nối bản ghi mới vào <code class="code-inline">users</code> và gọi <code class="code-inline">onAdded()</code> để reset <code class="code-inline">newUser = null</code>.
                    </li>
                </ul>

                <h3 class="report-h3">2.2. Tại sao state <code class="code-inline">kw</code> và <code class="code-inline">newUser</code> được đặt ở <code class="code-inline">App</code> mà không đặt ở <code class="code-inline">SearchForm</code> hay <code class="code-inline">AddUser</code>?</h3>
                <div class="callout-box callout-info">
                    <ol style="margin-left: 18px;">
                        <li style="margin-bottom: 8px;">
                            <strong>Giới hạn của Luồng dữ liệu một chiều (One-Way Data Flow):</strong> Trong React, <code class="code-inline">props</code> chỉ có thể truyền từ Component Cha xuống Component Con.
                            <code class="code-inline">SearchForm</code>, <code class="code-inline">AddUser</code> và <code class="code-inline">ResultTable</code> là các <em>Sibling Components</em> (anh em cùng cấp dưới <code class="code-inline">App</code>) nên <strong>không thể truyền trực tiếp state hay props sang cho nhau</strong>.
                        </li>
                        <li style="margin-bottom: 8px;">
                            <strong>Nếu đặt <code class="code-inline">kw</code> ở <code class="code-inline">SearchForm</code>:</strong> <code class="code-inline">SearchForm</code> chỉ biết giá trị người dùng gõ, nhưng component cần dùng từ khóa đó để lọc mảng (<code class="code-inline">users.filter(...)</code>) lại là <code class="code-inline">ResultTable</code>. Do đó <code class="code-inline">ResultTable</code> không thể đọc được <code class="code-inline">kw</code> nếu đặt trong <code class="code-inline">SearchForm</code>.
                        </li>
                        <li style="margin-bottom: 8px;">
                            <strong>Nếu đặt <code class="code-inline">newUser</code> ở <code class="code-inline">AddUser</code>:</strong> <code class="code-inline">AddUser</code> chỉ thu thập dữ liệu nhập từ form, còn mảng <code class="code-inline">users</code> hiển thị bảng lại nằm ở <code class="code-inline">ResultTable</code>. <code class="code-inline">ResultTable</code> sẽ không biết khi nào có người dùng mới được thêm để cập nhật bảng.
                        </li>
                        <li>
                            <strong>Giải pháp State Lifting tại <code class="code-inline">App</code>:</strong> Nâng cả hai state chia sẻ lên component cha chung gần nhất là <code class="code-inline">App</code> (<em>Single Source of Truth</em>). <code class="code-inline">App</code> truyền hàm setter xuống các con tạo dữ liệu và truyền giá trị state xuống con tiêu thụ. Ngược lại, các state chỉ dùng nội bộ giao diện (như <code class="code-inline">adding</code> hay <code class="code-inline">editing</code>) vẫn giữ ở component con.
                        </li>
                    </ol>
                </div>
            </section>

            <!-- PHẦN 3: MÔ TẢ PROPS -->
            <section class="report-section" id="sec-props">
                <h2 class="report-h2">3. Bảng mô tả Props truyền từ App xuống các Component con</h2>
                <p class="report-text">
                    Dưới đây là toàn bộ 5 Props được <code class="code-inline">App</code> truyền xuống 3 component con trong quá trình vận hành ứng dụng:
                </p>

                <div class="report-table-wrap">
                    <table>
                        <thead>
                            <tr>
                                <th>Component nhận</th>
                                <th>Tên Prop</th>
                                <th>Giá trị truyền từ App</th>
                                <th>Kiểu dữ liệu</th>
                                <th>Chiều dữ liệu</th>
                                <th>Mô tả vai trò &amp; cơ chế hoạt động</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td><strong>SearchForm</strong></td>
                                <td><span class="badge-prop">onChangeValue</span></td>
                                <td><code class="code-inline">setKeyword</code></td>
                                <td>Function</td>
                                <td><span class="badge-dir-up">Con &rarr; Cha (Callback)</span></td>
                                <td>Nhận hàm cập nhật state <code class="code-inline">kw</code>. Khi ô tìm kiếm phát sinh <code class="code-inline">onChange</code>, gọi <code class="code-inline">onChangeValue(e.target.value)</code> để cập nhật <code class="code-inline">kw</code> ở <code class="code-inline">App</code>.</td>
                            </tr>
                            <tr>
                                <td><strong>AddUser</strong></td>
                                <td><span class="badge-prop">onAdd</span></td>
                                <td><code class="code-inline">setNewUser</code></td>
                                <td>Function</td>
                                <td><span class="badge-dir-up">Con &rarr; Cha (Callback)</span></td>
                                <td>Nhận hàm cập nhật state <code class="code-inline">newUser</code>. Khi form hợp lệ và bấm <strong>Thêm</strong>, gọi <code class="code-inline">onAdd(user)</code> để đẩy object user mới lên <code class="code-inline">App</code>.</td>
                            </tr>
                            <tr>
                                <td><strong>ResultTable</strong></td>
                                <td><span class="badge-prop">keyword</span></td>
                                <td><code class="code-inline">kw</code></td>
                                <td>String</td>
                                <td><span class="badge-dir-down">Cha &rarr; Con (Data)</span></td>
                                <td>Truyền từ khóa tìm kiếm mới nhất từ <code class="code-inline">App</code> xuống để <code class="code-inline">ResultTable</code> lọc mảng <code class="code-inline">users</code> theo <code class="code-inline">name</code> và <code class="code-inline">username</code>.</td>
                            </tr>
                            <tr>
                                <td><strong>ResultTable</strong></td>
                                <td><span class="badge-prop">user</span></td>
                                <td><code class="code-inline">newUser</code></td>
                                <td>Object | null</td>
                                <td><span class="badge-dir-down">Cha &rarr; Con (Data)</span></td>
                                <td>Truyền dữ liệu người dùng mới từ <code class="code-inline">App</code> xuống. Hook <code class="code-inline">useEffect([user])</code> lắng nghe prop này để nối thêm bản ghi vào <code class="code-inline">users</code>.</td>
                            </tr>
                            <tr>
                                <td><strong>ResultTable</strong></td>
                                <td><span class="badge-prop">onAdded</span></td>
                                <td><code class="code-inline">() =&gt; setNewUser(null)</code></td>
                                <td>Function</td>
                                <td><span class="badge-dir-up">Con &rarr; Cha (Callback)</span></td>
                                <td>Callback báo lại <code class="code-inline">App</code> sau khi đã thêm xong người dùng vào bảng, giúp reset <code class="code-inline">newUser</code> về <code class="code-inline">null</code> để tránh thêm lặp ở lần render sau.</td>
                            </tr>
                        </tbody>
                    </table>
                </div>
            </section>

            <!-- PHẦN 4: LỖI & CÁCH KHẮC PHỤC -->
            <section class="report-section" id="sec-errors">
                <h2 class="report-h2">4. Các lỗi (Error / Warning) gặp phải, nguyên nhân và cách khắc phục</h2>
                <p class="report-text">
                    Trong quá trình cài đặt ứng dụng theo tài liệu hướng dẫn và kiểm thử CRUD, em đã gặp và xử lý 3 lỗi/cảnh báo tiêu biểu sau:
                </p>

                <!-- Lỗi 1 -->
                <div class="error-card err-runtime">
                    <div class="error-card-title">
                        <span class="tag-err tag-runtime">Runtime Error</span>
                        <span>Lỗi 1: <code class="code-inline">Too many re-renders</code> khi mở Modal Thêm người dùng ở <code class="code-inline">AddUser</code></span>
                    </div>
                    <p class="report-text"><strong>Thông báo lỗi trên Console:</strong></p>
                    <div class="error-console">Uncaught Error: Too many re-renders. React limits the number of renders to prevent an infinite loop.
    at renderWithHooks (react-dom.development.js:...)
    at AddUser (&lt;anonymous&gt;:...)</div>
                    <p class="report-text">
                        <strong>Nguyên nhân:</strong> Ở trang 6 của tài liệu hướng dẫn, nút Hủy trong <code class="code-inline">AddUser</code> được viết là
                        <code class="code-inline">&lt;button onClick={setAdding(false)}&gt;Hủy&lt;/button&gt;</code>.
                        Do viết trực tiếp dấu ngoặc đơn <code class="code-inline">()</code>, JavaScript thực thi ngay <code class="code-inline">setAdding(false)</code> trong lúc render
                        thay vì truyền vào hàm callback, gây vòng lặp cập nhật state &rarr; re-render vô tận.
                    </p>
                    <p class="report-text"><strong>Cách khắc phục:</strong> Truyền tham chiếu hàm hoặc bọc bằng Arrow Function để chỉ gọi khi người dùng click:</p>
                    <pre class="code-block">// Trước khi sửa (gọi trực tiếp khi render -&gt; crash):
&lt;button onClick={setAdding(false)}&gt;Hủy&lt;/button&gt;

// Sau khi sửa (truyền hàm callback hợp lệ):
&lt;button className="btn-secondary" onClick={handleCancel}&gt;Hủy&lt;/button&gt;
// hoặc: onClick={() =&gt; setAdding(false)}</pre>
                </div>

                <!-- Lỗi 2 -->
                <div class="error-card err-logic">
                    <div class="error-card-title">
                        <span class="tag-err tag-logic">Logic &amp; State Mutation</span>
                        <span>Lỗi 2: Sửa địa chỉ (<code class="code-inline">address</code>) trên Modal Sửa làm dữ liệu gốc trên bảng đổi theo ngay cả khi bấm Hủy</span>
                    </div>
                    <p class="report-text">
                        <strong>Hiện tượng sai:</strong> Khi nhấn nút <strong>Sửa</strong> ở <code class="code-inline">ResultTable</code> và gõ sửa <code class="code-inline">street</code>,
                        <code class="code-inline">suite</code> hoặc <code class="code-inline">city</code> trên Modal, dòng tương ứng trong bảng phía sau lập tức biến đổi theo từng phím gõ;
                        nếu nhấn <strong>Hủy</strong>, dữ liệu trên bảng vẫn bị sai lệch theo nội dung vừa gõ.
                    </p>
                    <p class="report-text">
                        <strong>Nguyên nhân:</strong> Nếu chỉ sao chép nông (<em>Shallow Copy</em>) bằng <code class="code-inline">setEditing({ ...user })</code>,
                        thuộc tính lồng nhau <code class="code-inline">editing.address</code> vẫn trỏ tới <strong>cùng vùng nhớ tham chiếu</strong> với <code class="code-inline">user.address</code>
                        trong state <code class="code-inline">users</code>, làm biến đổi trực tiếp state gốc của bảng.
                    </p>
                    <p class="report-text"><strong>Cách khắc phục:</strong> Dùng <strong>Deep Copy 2 tầng</strong> khi nạp state <code class="code-inline">editing</code> và khi cập nhật trường <code class="code-inline">address</code>:</p>
                    <pre class="code-block">function editUser(targetUser) {
    setEditing({
        ...targetUser,
        address: { ...(targetUser.address || { street: "", suite: "", city: "" }) }
    });
}

function handleEditChange(field, value) {
    if (["street", "suite", "city"].includes(field)) {
        setEditing((prev) =&gt; ({
            ...prev,
            address: { ...prev.address, [field]: value }
        }));
    } else {
        setEditing((prev) =&gt; ({ ...prev, [field]: value }));
    }
}</pre>
                </div>

                <!-- Lỗi 3 -->
                <div class="error-card err-warning">
                    <div class="error-card-title">
                        <span class="tag-err tag-warn">React Console Warning</span>
                        <span>Lỗi 3: Cảnh báo trùng lặp <code class="code-inline">key</code> khi Xóa rồi Thêm người dùng mới</span>
                    </div>
                    <p class="report-text"><strong>Thông báo Warning trên Console:</strong></p>
                    <div class="warn-console">Warning: Encountered two children with the same key, \`10\`. Keys should be unique so that components maintain their identity across updates.</div>
                    <p class="report-text">
                        <strong>Nguyên nhân:</strong> Nếu gán ID mới bằng công thức <code class="code-inline">id: prev.length + 1</code> (theo mã gợi ý cơ bản):
                        danh sách ban đầu có 10 phần tử (<code class="code-inline">id</code> 1..10). Khi xóa 1 user ở giữa (ví dụ <code class="code-inline">id=3</code>),
                        mảng còn 9 phần tử nhưng vẫn chứa phần tử <code class="code-inline">id=10</code>; khi thêm mới thì <code class="code-inline">9 + 1 = 10</code>,
                        tạo ra hai dòng <code class="code-inline">&lt;tr key={u.id}&gt;</code> trùng <code class="code-inline">key="10"</code>.
                    </p>
                    <p class="report-text"><strong>Cách khắc phục:</strong> Sinh ID mới theo giá trị ID lớn nhất hiện có trong mảng kết hợp truyền đầy đủ <code class="code-inline">key={u.id}</code>:</p>
                    <pre class="code-block">const nextId = prev.length &gt; 0 ? Math.max(...prev.map((u) =&gt; u.id)) + 1 : 1;
return [...prev, { ...user, id: nextId }];</pre>
                </div>
            </section>

            <!-- PHẦN 5: HOST & KIỂM THỬ -->
            <section class="report-section" id="sec-host">
                <h2 class="report-h2">5. Kết quả Triển khai Host (Vercel) &amp; Kiểm thử CRUD</h2>
                <div class="report-table-wrap">
                    <table>
                        <thead>
                            <tr>
                                <th style="width: 34%;">Hạng mục</th>
                                <th>Liên kết trực tiếp / Chi tiết</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td><strong>Ứng dụng React CRUD + Báo cáo (Live)</strong></td>
                                <td><a class="report-link" href="https://it-4409.vercel.app/Bai%20tap%20React%20-%20CRUD/index.html" target="_blank" rel="noopener">https://it-4409.vercel.app/Bai%20tap%20React%20-%20CRUD/index.html</a></td>
                            </tr>
                            <tr>
                                <td><strong>Trang Báo cáo độc lập (Live)</strong></td>
                                <td><a class="report-link" href="https://it-4409.vercel.app/Bai%20tap%20React%20-%20CRUD/bao-cao.html" target="_blank" rel="noopener">https://it-4409.vercel.app/Bai%20tap%20React%20-%20CRUD/bao-cao.html</a></td>
                            </tr>
                            <tr>
                                <td><strong>Tải Báo cáo Word trên Host (.docx)</strong></td>
                                <td><a class="report-link" href="https://it-4409.vercel.app/Bai%20tap%20React%20-%20CRUD/Bao_cao_Bai_tap_React_CRUD.docx" target="_blank" rel="noopener">https://it-4409.vercel.app/Bai%20tap%20React%20-%20CRUD/Bao_cao_Bai_tap_React_CRUD.docx</a></td>
                            </tr>
                            <tr>
                                <td><strong>Trang chủ IT4409</strong></td>
                                <td><a class="report-link" href="https://it-4409.vercel.app/" target="_blank" rel="noopener">https://it-4409.vercel.app/</a></td>
                            </tr>
                            <tr>
                                <td><strong>GitHub Repository</strong></td>
                                <td><a class="report-link" href="https://github.com/vuhoabinh2416137/IT4409" target="_blank" rel="noopener">https://github.com/vuhoabinh2416137/IT4409</a></td>
                            </tr>
                        </tbody>
                    </table>
                </div>

                <div class="report-table-wrap">
                    <table>
                        <thead>
                            <tr>
                                <th style="width: 24%;">Chức năng</th>
                                <th style="width: 30%;">Thao tác thực hiện</th>
                                <th>Cơ chế xử lý trong ứng dụng</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td><strong>Read (Hiển thị)</strong></td>
                                <td>Mở trang lần đầu</td>
                                <td><code class="code-inline">useEffect(..., [])</code> gọi Fetch API JSONPlaceholder và render bảng 10 user ban đầu.</td>
                            </tr>
                            <tr>
                                <td><strong>Search (Tìm kiếm)</strong></td>
                                <td>Gõ từ khóa vào <code class="code-inline">SearchForm</code></td>
                                <td>Nâng từ khóa lên <code class="code-inline">App (kw)</code>, truyền xuống <code class="code-inline">ResultTable (keyword)</code> để lọc theo <code class="code-inline">name</code> &amp; <code class="code-inline">username</code>.</td>
                            </tr>
                            <tr>
                                <td><strong>Create (Thêm mới)</strong></td>
                                <td>Bấm "+ Thêm người dùng", nhập &amp; Thêm</td>
                                <td>Validate Name/Username, gửi <code class="code-inline">user</code> lên <code class="code-inline">App (newUser)</code>, <code class="code-inline">ResultTable</code> thêm với ID duy nhất và gọi <code class="code-inline">onAdded()</code>.</td>
                            </tr>
                            <tr>
                                <td><strong>Update (Chỉnh sửa)</strong></td>
                                <td>Bấm "Sửa", chỉnh thông tin &amp; Lưu</td>
                                <td>Deep Copy vào <code class="code-inline">editing</code>; Hủy giữ nguyên bảng gốc; Lưu thay thế bằng <code class="code-inline">prev.map(...)</code>.</td>
                            </tr>
                            <tr>
                                <td><strong>Delete (Xóa)</strong></td>
                                <td>Bấm "Xóa" trên một dòng</td>
                                <td>Cập nhật mảng bất biến với <code class="code-inline">setUsers(prev =&gt; prev.filter(u =&gt; u.id !== id))</code>.</td>
                            </tr>
                        </tbody>
                    </table>
                </div>
            </section>
        `;
    }

    window.renderReactCrudReport = renderReactCrudReport;
})();
