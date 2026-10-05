# Xuất hợp đồng

Công cụ tạo hợp đồng lao động và hợp đồng dịch vụ từ một file danh sách Excel:
nạp file, chọn người, xuất ra Word hoặc PDF.

Có hai cách dùng, chung một lõi xử lý:

| | Chạy ở đâu | Dùng khi nào |
|---|---|---|
| **Bản web** | Ngay trong trình duyệt (Pyodide) | Mở link là dùng, không cần cài gì |
| **Bản trên máy** | Máy chủ FastAPI chạy tại máy bạn | Khi muốn nhanh hơn và dùng được offline |

## Dữ liệu không nằm trong kho mã này

File danh sách `Onboarding_List.xlsx`, thư mục `CONTRACTS/output/` và script
`CONTRACTS/scripts/build_excel.py` đều bị `.gitignore` chặn và **không bao giờ**
được đẩy lên. Bản web cũng không gửi file của bạn đi đâu: Python chạy ngay trong
tab trình duyệt, file Excel chỉ được đọc trong bộ nhớ máy bạn.

## Bản web

Trang tĩnh ở thư mục gốc (`index.html` + `web/backend-pyodide.js`). GitHub Pages
phục vụ thẳng từ nhánh `main`. Lần mở đầu tiên mất khoảng 10–30 giây để tải
Python và các thư viện về; sau đó trình duyệt lưu đệm nên vào lại là chạy ngay.

Thử tại máy trước khi đẩy lên:

```bash
python3 -m http.server 8899     # rồi mở http://127.0.0.1:8899
```

## Bản trên máy

```bash
python3 -m venv .venv && .venv/bin/pip install -r requirements.txt
./CONTRACTS/start_web.command    # mở http://127.0.0.1:8777
```

## Đẩy thay đổi lên

`git push` lên nhánh `main` là xong — GitHub Pages tự cập nhật trang sau khoảng
một phút, không cần build gì.

Mỗi lần đẩy, GitHub chạy `CONTRACTS/scripts/kiem_tra.py` để soi năm thứ:

1. Không file nào chứa dữ liệu cá nhân (Excel, hợp đồng đã xuất, CCCD, email, số điện thoại)
2. Bản web khai đủ file cần tải — thêm module Python mới mà quên khai vào
   `web/backend-pyodide.js` thì bản trên máy vẫn chạy còn bản web chết ngay lúc mở
3. Hai trang `index.html` đều có đủ phần tử mà `app.js` cần
4. Dựng thử cả hai mẫu hợp đồng ra Word và PDF, và đối chiếu biến trong template
   với dữ liệu code thật sự cấp
5. Cú pháp JavaScript

Chạy trước ở máy cho nhanh:

```bash
.venv/bin/python CONTRACTS/scripts/kiem_tra.py
```

## Sửa nội dung hợp đồng

**Hai file `.docx` trong `CONTRACTS/templates/` là bản gốc duy nhất.** Mở bằng
Word, sửa chữ nghĩa và định dạng (thụt lề, cỡ chữ, đánh số, giãn dòng, bảng
biểu) rồi lưu lại — cả Word lẫn PDF xuất ra đều đi theo đúng file đó. Chỗ điền
dữ liệu viết theo cú pháp `{{ ten_bien }}`.

`CONTRACTS/app/contract_spec.py` và `scripts/build_templates.py` chỉ dùng để dựng
template lần đầu. Chạy `build_templates.py --force` sẽ ghi đè công sức bạn sửa
trong Word (có sao lưu sang `templates/backup/`), nên bình thường đừng đụng tới.

## Cấu trúc

```
index.html                     trang cho GitHub Pages
web/backend-pyodide.js         cầu nối JS ↔ Python khi chạy trong trình duyệt
CONTRACTS/
  app/
    excel.py                   đọc file danh sách thành dữ liệu hợp đồng
    render_docx.py             điền dữ liệu vào template Word
    docx_read.py               đọc định dạng trong file .docx đã điền
    render_pdf.py              dựng PDF theo đúng định dạng đọc được
    browser.py                 đầu vào cho bản chạy trong trình duyệt
    server.py                  đầu vào cho bản chạy trên máy (FastAPI)
    static/                    giao diện, dùng chung cho cả hai bản
  templates/                   bản gốc hợp đồng (sửa trong Word)
  fonts/                       Tinos — font thay Times New Roman khi chạy web
```

Font [Tinos](https://fonts.google.com/specimen/Tinos) trùng khít số đo với Times
New Roman (cùng chiều cao chữ và bề rộng từng ký tự, kể cả dấu tiếng Việt) nên
PDF dựng trong trình duyệt xuống dòng và sang trang y hệt bản dựng trên máy.
Giấy phép SIL OFL, xem `CONTRACTS/fonts/OFL.txt`.
