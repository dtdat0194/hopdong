# -*- coding: utf-8 -*-
"""Khung nội dung hai mẫu hợp đồng, chỉ dùng để dựng template Word lần đầu.

BẢN GỐC LÀ FILE .docx TRONG templates/ - sửa nội dung và định dạng trực tiếp trong
Word, PDF sẽ xuất theo đúng file đó. File này KHÔNG còn là nguồn của bản xuất ra;
chạy lại scripts/build_templates.py --force sẽ ghi đè công sức bạn đã sửa trong Word.

Phần CONTRACTS ở cuối file vẫn được dùng thường xuyên: nó khai báo tên hợp đồng,
tên file xuất ra và template tương ứng cho mỗi loại.
"""

FONT = "Times New Roman"
SIZE = 13
PAGE = dict(width=21.0, height=29.7, top=2.0, bottom=2.0, left=3.0, right=2.0)  # cm


def P(text="", **kw):
    return dict(k="p", text=text, **kw)


def HEAD(text):
    return dict(k="p", text=text, bold=True, before=6, keep=True)


def NUM(n, text, **kw):
    kw.setdefault("left", 0.75)
    kw.setdefault("hanging", 0.75)
    return dict(k="p", text=f"{n} {text}", **kw)


def SUB(text, **kw):
    kw.setdefault("left", 1.5)
    kw.setdefault("after", 0)
    return dict(k="p", text=text, **kw)


def IND(text, **kw):
    kw.setdefault("left", 0.75)
    return dict(k="p", text=text, **kw)


def TABLE(rows, label_w=4.2):
    return dict(k="table", rows=rows, label_w=label_w)


def SIGN():
    return dict(k="sign")


def TITLE(title, so_hieu):
    return [
        P("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", bold=True, align="center", after=0),
        P("Độc lập - Tự do - Hạnh phúc", bold=True, align="center", underline=True, after=12),
        P(title, bold=True, align="center", size=15, after=0),
        P(so_hieu, bold=True, align="center", after=12),
    ]


BEN_A = [
    ("BÊN A", "{{ cong_ty }}"),
    ("Địa chỉ trụ sở", "{{ dia_chi_cong_ty }}"),
    ("Điện thoại", "{{ dien_thoai_cong_ty }}"),
    ("Mã số thuế", "{{ mst_cong_ty }}"),
    ("Đại diện bởi", "{{ xung_ho_dai_dien }} {{ dai_dien }}"),
    ("Chức vụ", "{{ chuc_vu_dai_dien }}"),
]

BEN_B = [
    ("BÊN B", "{{ ho_ten }}"),
    ("Ngày sinh", "{{ ngay_sinh }}"),
    ("Nơi cư trú", "{{ noi_cu_tru }}"),
    ("Số CCCD", "{{ so_cccd }}"),
    ("Ngày cấp", "{{ ngay_cap_cccd }}"),
    ("Nơi cấp", "{{ noi_cap_cccd }}"),
]


# ------------------------------------------------------------------ HỢP ĐỒNG DỊCH VỤ KỸ SƯ
def hop_dong_dich_vu():
    b = TITLE("HỢP ĐỒNG DỊCH VỤ KỸ SƯ", "Số: {{ so_hop_dong }}.{{ nam_hop_dong }}/{{ ky_hieu_hd }}")
    b += [
        P("Hợp đồng Dịch vụ kỹ sư (“Hợp Đồng”) này được lập vào ngày {{ ngay_ky }} tháng "
          "{{ thang_ky }} năm {{ nam_ky }}, giữa và bởi:", keep=True),
        TABLE(BEN_A),
        P("(Sau đây gọi là “Công ty”/“Bên A”)", italic=True, before=4),
        P("Và"),
        TABLE(BEN_B),
        P("(Sau đây gọi là “Bên B”)", italic=True, before=4, after=8),

        P("XÉT RẰNG:", bold=True),
        P("- Bên A có nhu cầu hợp tác với cá nhân có chuyên môn phù hợp để thực hiện các công việc "
          "thuộc lĩnh vực {{ linh_vuc_cong_viec }} nhằm phục vụ hoạt động nghiên cứu, phát triển "
          "sản phẩm và vận hành của Công ty."),
        P("- Bên B là cá nhân có kinh nghiệm, chuyên môn phù hợp và đồng ý hợp tác thực hiện Công Việc "
          "theo hình thức độc lập, không thuộc biên chế lao động của Bên A."),
        P("Sau khi thỏa thuận, Các Bên thống nhất ký kết Hợp Đồng với các điều khoản và điều kiện sau:",
          after=8),

        HEAD("Điều 1: Phạm vi công việc"),
        NUM("1.", "Bên B thực hiện các đầu việc (“Công Việc”) theo thỏa thuận giữa hai Bên vào đầu mỗi tháng, "
                  "cụ thể tại Bảng phân công công việc trao đổi bằng văn bản hoặc email giữa hai Bên."),
        NUM("2.", "Bên B chủ động sắp xếp thời gian, phương thức và địa điểm thực hiện Công Việc được Bên A "
                  "giao nhằm đảm bảo tiến độ và chất lượng Công Việc."),
        NUM("3.", "Bên B báo cáo tiến độ Công Việc theo các mốc thống nhất giữa hai Bên."),
        NUM("4.", "Cuối mỗi tháng, hai Bên lập Biên bản nghiệm thu khối lượng và kết quả Công Việc hoàn thành, "
                  "làm căn cứ thanh toán Thù lao theo Điều 4."),

        HEAD("Điều 2: Quyền và nghĩa vụ của Bên A"),
        NUM("1.", "Giao đầu việc, cung cấp thông tin, tài liệu và thiết bị cần thiết cho Bên B, tạo điều kiện "
                  "thuận lợi cho Bên B trong quá trình thực hiện công việc."),
        NUM("2.", "Xác nhận/nghiệm thu kết quả Công Việc hằng tháng đúng thời hạn thỏa thuận."),
        NUM("3.", "Thanh toán đầy đủ, đúng hạn Thù lao cho Bên B theo Điều 4."),

        HEAD("Điều 3: Quyền và nghĩa vụ của Bên B"),
        NUM("1.", "Chủ động thực hiện Công Việc, tự chịu trách nhiệm về phương thức và tiến độ thực hiện, "
                  "đảm bảo hoàn thành đúng kết quả đã thống nhất với Bên A."),
        NUM("2.", "Báo cáo tiến độ Công Việc theo các mốc đã thống nhất."),
        NUM("3.", "Không được chuyển giao Công Việc cho bên thứ ba thực hiện khi chưa có sự đồng ý bằng văn bản "
                  "của Bên A."),
        NUM("4.", "Bên B chịu trách nhiệm bồi thường cho Bên A đối với các thiệt hại trực tiếp phát sinh do lỗi "
                  "của Bên B trong quá trình thực hiện Công Việc. Bên B không chịu trách nhiệm đối với khách hàng "
                  "hoặc bên thứ ba, trừ trường hợp thiệt hại phát sinh trực tiếp do lỗi cố ý hoặc vi phạm nghiêm "
                  "trọng Hợp đồng của Bên B. Tổng trách nhiệm bồi thường của Bên B không vượt quá tổng giá trị "
                  "Hợp đồng."),

        HEAD("Điều 4: Thù lao và thanh toán"),
        NUM("1.", "Mức phí: {{ muc_luong }} VNĐ ({{ muc_luong_bang_chu }}), chưa khấu trừ thuế thu nhập cá nhân, "
                  "áp dụng cho khối lượng Công Việc đã thống nhất tại Điều 1 và gắn với kết quả xác nhận tại "
                  "Biên bản nghiệm thu hằng tháng."),
        IND("Đối với mỗi lần chi trả, Bên A khấu trừ {{ ty_le_khau_tru }}% thuế thu nhập cá nhân trước khi "
            "thanh toán theo quy định pháp luật hiện hành, trừ trường hợp Bên B đủ điều kiện và có cam kết "
            "tạm thời không khấu trừ theo mẫu quy định."),
        NUM("2.", "Thù lao được thanh toán theo từng tháng, trong vòng 05 ngày kể từ ngày hai Bên ký Biên bản "
                  "nghiệm thu/xác nhận khối lượng công việc của tháng đó."),
        NUM("3.", "Trường hợp khối lượng Công Việc thực tế thay đổi đáng kể so với thỏa thuận ban đầu, hai Bên "
                  "thống nhất điều chỉnh mức phí bằng Phụ lục Hợp đồng."),
        NUM("4.", "Thanh toán được thực hiện bằng chuyển khoản vào tài khoản:", keep=True),
        SUB("Số tài khoản ngân hàng: {{ so_tai_khoan }}", keep=True),
        SUB("Tên ngân hàng: {{ ten_ngan_hang }}", keep=True),
        SUB("Chi nhánh: {{ chi_nhanh }}", cond="chi_nhanh", keep=True),
        SUB("Chủ tài khoản: {{ chu_tai_khoan }}", after=4),

        HEAD("Điều 5: Tuyên bố và cam kết về trách nhiệm của hai bên"),
        NUM("1.", "Bên B hiểu và đồng ý rằng việc cung cấp Công Việc của Bên B theo Hợp đồng này sẽ chỉ với tư cách "
                  "là một nhà thầu độc lập mà không cấu thành quan hệ lao động, đại lý hoặc bất kỳ quan hệ nào khác "
                  "với Bên A. Bên B không có quyền ràng buộc Bên A bằng bất kỳ hợp đồng hoặc thỏa thuận nào khác về "
                  "phạm vi Công Việc được cung cấp theo Hợp đồng này."),
        NUM("2.", "Bên B cam kết thực hiện Công Việc đúng nhiệm vụ và phạm vi được giao. Bên B chỉ chịu trách nhiệm "
                  "đối với khiếu nại, khiếu kiện của bên thứ ba nếu phát sinh trực tiếp do lỗi cố ý hoặc vi phạm "
                  "nghiêm trọng Hợp đồng của Bên B. Quy định này không áp dụng đối với kết quả Công Việc đã được "
                  "Bên A nghiệm thu, phê duyệt, trừ trường hợp Bên B cố ý che giấu lỗi hoặc cung cấp thông tin sai "
                  "lệch tại thời điểm nghiệm thu. Bên A có trách nhiệm rà soát và phê duyệt trước khi áp dụng hoặc "
                  "bàn giao cho bên thứ ba. Tổng trách nhiệm của Bên B thực hiện theo Điều 8.2 của Hợp đồng."),
        NUM("3.", "Hợp đồng này không làm phát sinh nghĩa vụ của Bên A trong việc đóng bảo hiểm xã hội, bảo hiểm y tế, "
                  "bảo hiểm thất nghiệp bắt buộc cho Bên B."),

        HEAD("Điều 6: Quyền sở hữu trí tuệ"),
        NUM("1.", "Bên A là chủ sở hữu quyền tài sản đối với toàn bộ kết quả Công Việc (bao gồm nhưng không giới hạn: "
                  "mã nguồn, thuật toán, tài liệu kỹ thuật, thiết kế, dữ liệu, mô hình...) do Bên B tạo ra trong quá "
                  "trình thực hiện Hợp đồng này ngay khi được tạo ra hoặc bàn giao, tùy thời điểm nào đến trước."),
        NUM("2.", "Bên B cam kết kết quả Công Việc là sản phẩm sáng tạo của chính mình, không sao chép hay vi phạm "
                  "quyền sở hữu trí tuệ của bất kỳ bên thứ ba nào; chịu trách nhiệm giải quyết và bồi thường thiệt hại "
                  "nếu phát sinh tranh chấp về quyền sở hữu trí tuệ liên quan đến kết quả Công Việc do mình cung cấp."),
        NUM("3.", "Bên B không được sử dụng tài sản, dữ liệu, tài liệu do Bên A cung cấp cho bất kỳ mục đích nào khác "
                  "ngoài phạm vi Hợp đồng này khi chưa có sự đồng ý bằng văn bản của Bên A."),

        HEAD("Điều 7: Bảo mật thông tin"),
        NUM("1.", "Bên B có trách nhiệm giữ bí mật toàn bộ thông tin, tài liệu liên quan đến Bên A mà Bên B biết được "
                  "trong quá trình thực hiện Hợp đồng này, không tiết lộ cho bên thứ ba khi chưa có sự đồng ý bằng văn "
                  "bản của Bên A, trừ trường hợp phục vụ mục đích thực hiện Hợp đồng."),
        NUM("2.", "Nghĩa vụ bảo mật tại Điều này tiếp tục có hiệu lực trong vòng 05 (năm) năm kể từ ngày Hợp đồng "
                  "chấm dứt."),

        HEAD("Điều 8: Phạt vi phạm và bồi thường thiệt hại"),
        NUM("1.", "Trường hợp một Bên vi phạm nghĩa vụ theo Hợp đồng này, Bên bị vi phạm có quyền yêu cầu Bên vi phạm "
                  "thanh toán khoản tiền phạt vi phạm tương đương 10% giá trị phần nghĩa vụ bị vi phạm, cùng với bồi "
                  "thường thiệt hại thực tế phát sinh (nếu có)."),
        NUM("2.", "Một phần và/hoặc toàn bộ trách nhiệm của mỗi Bên theo Hợp Đồng này đối với bất kỳ việc khiếu nại, "
                  "tranh chấp phát sinh từ hoặc có liên quan đến Hợp Đồng này, trong mọi trường hợp, sẽ không vượt quá "
                  "tổng giá trị Hợp Đồng, trừ trường hợp vi phạm Điều 6 và Điều 7."),

        HEAD("Điều 9: Thời hạn hợp đồng"),
        NUM("1.", "Thời hạn thực hiện Hợp đồng từ ngày {{ ngay_bat_dau }} đến hết ngày {{ ngay_ket_thuc }} "
                  "({{ so_thang }} tháng)."),
        NUM("2.", "Trước khi hết thời hạn, nếu hai Bên có nhu cầu tiếp tục hợp tác, sẽ ký Phụ lục gia hạn Hợp đồng "
                  "nêu rõ thời hạn mới và mức phí áp dụng (nếu thay đổi)."),

        HEAD("Điều 10: Giải quyết tranh chấp"),
        NUM("1.", "Mọi tranh chấp phát sinh từ hoặc liên quan đến Hợp đồng này trước hết được hai Bên giải quyết "
                  "thông qua thương lượng trên tinh thần thiện chí."),
        NUM("2.", "Trường hợp thương lượng không thành trong thời hạn 30 ngày kể từ ngày một Bên gửi yêu cầu, mỗi Bên "
                  "có quyền yêu cầu Tòa án có thẩm quyền giải quyết theo quy định của pháp luật."),
        IND("Bên A có quyền tạm giữ các khoản tiền phải thanh toán cho Bên B cho đến khi các Bên giải quyết xong các "
            "tranh chấp từ hoặc liên quan đến Hợp Đồng này."),

        HEAD("Điều 11: Chấm dứt hợp đồng"),
        P("Hợp đồng chấm dứt trong các trường hợp sau:"),
        NUM("1.", "Hết thời hạn quy định tại Điều 9 mà hai Bên không gia hạn;"),
        NUM("2.", "Theo thỏa thuận bằng văn bản của hai Bên;"),
        NUM("3.", "Một Bên vi phạm nghĩa vụ và không khắc phục trong 10 (mười) ngày kể từ ngày nhận được văn bản "
                  "yêu cầu của Bên kia;"),
        NUM("4.", "Một Bên đơn phương chấm dứt bằng thông báo văn bản gửi Bên kia trước 15 (mười lăm) ngày. Hợp đồng "
                  "sẽ tự động chấm dứt nếu hai Bên không vi phạm các nội dung của Điều 6 và Điều 7."),

        HEAD("Điều 12: Điều khoản thi hành"),
        NUM("1.", "Hợp đồng có hiệu lực kể từ ngày ký, áp dụng đối với Công Việc Bên B thực hiện từ ngày "
                  "{{ ngay_bat_dau }}, và được điều chỉnh theo pháp luật Việt Nam."),
        NUM("2.", "Mọi sửa đổi, bổ sung Hợp đồng phải được lập thành văn bản có chữ ký của hai Bên."),
        NUM("3.", "Hợp đồng được lập thành 02 (hai) bản tiếng Việt, có giá trị pháp lý như nhau, mỗi Bên giữ "
                  "01 (một) bản.", keep=True),
        SIGN(),
    ]
    return b


# ------------------------------------------------------------------ HỢP ĐỒNG LAO ĐỘNG
def hop_dong_lao_dong():
    ben_a = [("BÊN A (Người sử dụng lao động)", "{{ cong_ty }}")] + BEN_A[1:]
    ben_a[3] = ("Mã số thuế", "{{ mst_cong_ty }} do {{ noi_cap_dkkd }} cấp ngày {{ ngay_cap_dkkd }}.")
    # HĐ lao động giữ nhãn "Sinh ngày"; HĐ dịch vụ dùng "Ngày sinh".
    ben_b = [("BÊN B (Người lao động)", "{{ ho_ten }}"), ("Sinh ngày", "{{ ngay_sinh }}")] + BEN_B[2:]

    b = TITLE("HỢP ĐỒNG LAO ĐỘNG", "Số: {{ so_hop_dong }}.{{ nam_hop_dong }}/{{ ky_hieu_hd }}")
    b += [
        P("Căn cứ vào Bộ Luật Lao Động số 45/2019/QH14 ngày 20/11/2019;", italic=True),
        P("Căn cứ vào Nghị Định số 145/2020/NĐ-CP ngày 14/12/2020 của Chính Phủ hướng dẫn chi tiết một số "
          "nội dung của Bộ Luật Lao Động.", italic=True),
        P("Hôm nay, ngày {{ ngay_ky }} tháng {{ thang_ky }} năm {{ nam_ky }}, chúng tôi gồm:", after=6, keep=True),
        TABLE(ben_a, label_w=6.0),
        P("(Sau đây gọi là “Công ty”/“Bên A”)", italic=True, before=4),
        P("Và"),
        TABLE(ben_b, label_w=6.0),
        P("(Sau đây gọi là “Bên B”)", italic=True, before=4),
        P("Hai Bên thỏa thuận ký kết Hợp đồng lao động và cam kết đúng các điều khoản sau:", after=8),

        HEAD("Điều 1: Công việc và địa điểm làm việc"),
        IND("Chức danh: {{ chuc_danh }}", after=0),
        IND("Bộ phận: {{ bo_phan }}", after=0),
        IND("Địa điểm làm việc: {{ dia_diem_lam_viec }}"),

        HEAD("Điều 2: Loại hợp đồng và thời hạn"),
        IND("Loại hợp đồng: Hợp đồng lao động xác định thời hạn {{ so_thang }} tháng", after=0),
        IND("Thời gian thực hiện: từ ngày {{ ngay_bat_dau }} đến ngày {{ ngay_ket_thuc }}"),

        HEAD("Điều 3: Thời gian làm việc và ngày nghỉ"),
        NUM("1.", "Số giờ làm việc: {{ gio_tuan }} giờ/tuần theo lịch làm việc chung của Công ty."),
        NUM("2.", "Lịch làm việc cụ thể:", keep=True),
        SUB("a. Trong ngày: {{ gio_ngay }} giờ/ngày - Sáng từ 8h00 đến 12h00, Chiều từ 13h00 đến 17h00",
            keep=True),
        SUB("b. Trong tuần: {{ gio_tuan }} giờ/tuần - {{ so_ngay_tuan }} ngày/tuần từ thứ Hai đến thứ Sáu",
            after=4),
        NUM("3.", "Nghỉ hằng tuần: Thứ Bảy và Chủ nhật."),
        NUM("4.", "Nghỉ lễ, tết: theo quy định của pháp luật lao động; lịch nghỉ cụ thể do Công ty thông báo "
                  "hằng năm."),
        NUM("5.", "Nghỉ phép năm hưởng nguyên lương: {{ ngay_phep_nam }} ngày/năm; đủ 05 năm làm việc cho Công ty "
                  "được tăng thêm tương ứng theo quy định của pháp luật."),

        HEAD("Điều 4: Quyền lợi và nghĩa vụ của người lao động"),
        IND("Quyền lợi:", bold=True, keep=True),
        NUM("1.", "Mức lương: {{ muc_luong }} VNĐ/tháng (Bằng chữ: {{ muc_luong_bang_chu }}).", keep=True),
        IND("Đây là mức lương gộp, đã bao gồm các khoản bảo hiểm bắt buộc thuộc trách nhiệm của Bên B và thuế "
            "thu nhập cá nhân (TNCN)."),
        NUM("2.", "Phụ cấp và các khoản bổ sung khác: Theo chính sách của Công ty."),
        NUM("3.", "Bên B được tham gia BHXH, BHYT, BHTN, bảo hiểm tai nạn lao động và bệnh nghề nghiệp kể từ tháng "
                  "Hợp đồng lao động này có hiệu lực. Mức đóng và tỷ lệ đóng theo quy định pháp luật và được điều "
                  "chỉnh khi pháp luật thay đổi."),
        NUM("4.", "Hình thức trả lương: Chuyển khoản vào tài khoản ngân hàng của Bên B:", keep=True),
        SUB("a. Số tài khoản: {{ so_tai_khoan }}", keep=True),
        SUB("b. Ngân hàng: {{ ten_ngan_hang }}{% if chi_nhanh %}, {{ chi_nhanh }}{% endif %}", keep=True),
        SUB("c. Chủ tài khoản: {{ chu_tai_khoan }}", after=4),
        NUM("5.", "Kỳ hạn trả lương: Trả 01 lần/tháng, vào ngày {{ ngay_tra_luong }} của tháng liền kề sau tháng "
                  "làm việc. Nếu ngày {{ ngay_tra_luong }} trùng ngày nghỉ hằng tuần hoặc ngày nghỉ lễ, tết thì "
                  "lương được trả vào ngày làm việc liền sau đó."),
        NUM("6.", "Các khoản khấu trừ: Trước khi chi trả, Công ty khấu trừ phần BHXH, BHYT, BHTN thuộc trách nhiệm "
                  "của Bên B và thuế TNCN theo quy định pháp luật."),
        NUM("7.", "Tháng lương thứ 13: Theo chính sách của Công ty."),
        NUM("8.", "Tiền thưởng: Theo quy chế khen thưởng của Công ty."),
        NUM("9.", "Chế độ nâng lương: Theo quy chế điều chỉnh lương của Công ty."),
        IND("Nghĩa vụ:", bold=True, keep=True),
        IND("Thực hiện công việc theo Hợp đồng dưới sự quản lý, điều hành, giám sát của Công ty; chấp hành nội quy "
            "lao động, quy định về an toàn, vệ sinh lao động."),

        HEAD("Điều 5: Quyền và nghĩa vụ của Người sử dụng lao động"),
        NUM("1.", "Bố trí, điều hành Người lao động hoàn thành công việc theo Hợp đồng; ban hành và áp dụng nội quy "
                  "lao động đối với Người lao động."),
        NUM("2.", "Thực hiện đầy đủ nghĩa vụ tiền lương, bảo hiểm và các chế độ khác theo Hợp đồng và quy định "
                  "pháp luật."),

        HEAD("Điều 6: Trang bị bảo hộ lao động"),
        NUM("1.", "Công ty có trách nhiệm trang bị đầy đủ phương tiện bảo hộ lao động cần thiết cho Người lao động "
                  "theo quy định kèm biên bản bàn giao. Người lao động có trách nhiệm sử dụng đúng và đầy đủ các "
                  "phương tiện được cấp trong quá trình làm việc."),
        NUM("2.", "Người lao động có trách nhiệm bảo quản, sử dụng đúng mục đích công việc các công cụ, thiết bị, "
                  "tài sản được Công ty giao, và bàn giao lại đầy đủ khi chấm dứt Hợp đồng hoặc khi Công ty yêu cầu."),
        NUM("3.", "Trường hợp làm hư hỏng, mất công cụ, thiết bị, tài sản của Công ty do lỗi của Người lao động, "
                  "Người lao động có trách nhiệm bồi thường theo quy định tại Điều 129 Bộ luật Lao động 2019 và "
                  "nội quy lao động của Công ty."),

        HEAD("Điều 7: Đào tạo, bồi dưỡng, nâng cao trình độ, kỹ năng nghề"),
        IND("Công ty tạo điều kiện đào tạo, bồi dưỡng nâng cao trình độ, kỹ năng nghề cho Người lao động theo nhu cầu "
            "công việc và chính sách của Công ty."),

        HEAD("Điều 8: Bảo mật thông tin và quyền sở hữu trí tuệ"),
        NUM("1.", "Người lao động có nghĩa vụ bảo mật thông tin của Công ty, khách hàng và đối tác theo Thỏa thuận "
                  "bảo mật thông tin."),
        NUM("2.", "Quyền sở hữu trí tuệ đối với sản phẩm do Người lao động tạo ra khi thực hiện công việc được giao "
                  "được xác định theo Thỏa thuận chuyển giao quyền sở hữu trí tuệ."),
        NUM("3.", "Hai thỏa thuận trên được ký kèm, là bộ phận không tách rời của Hợp đồng này và tiếp tục có hiệu "
                  "lực sau khi Hợp đồng chấm dứt."),

        HEAD("Điều 9: Giải quyết tranh chấp"),
        NUM("1.", "Mọi tranh chấp phát sinh từ hoặc liên quan đến Hợp đồng này trước hết được hai Bên giải quyết "
                  "thông qua thương lượng trên tinh thần thiện chí."),
        NUM("2.", "Trường hợp thương lượng không thành trong thời hạn 30 ngày kể từ ngày một Bên gửi yêu cầu thương "
                  "lượng, giải quyết tranh chấp theo trình tự, thủ tục quy định tại Bộ luật Lao động 2019 và pháp "
                  "luật có liên quan."),

        HEAD("Điều 10: Chấm dứt hợp đồng lao động"),
        IND("Việc đơn phương chấm dứt Hợp đồng của mỗi Bên thực hiện theo quy định tại Điều 34, 35, 36 Bộ luật Lao "
            "động 2019, bao gồm điều kiện, lý do và thời hạn báo trước tương ứng."),

        HEAD("Điều 11: Điều khoản thi hành"),
        NUM("1.", "Hợp đồng có hiệu lực kể từ ngày ký, áp dụng đối với công việc Bên B thực hiện từ ngày "
                  "{{ ngay_bat_dau }}, và được điều chỉnh theo pháp luật Việt Nam."),
        NUM("2.", "Mọi sửa đổi, bổ sung Hợp đồng phải được lập thành văn bản có chữ ký của hai Bên."),
        NUM("3.", "Hợp đồng được lập thành 02 (hai) bản tiếng Việt, có giá trị pháp lý như nhau, mỗi Bên giữ "
                  "01 (một) bản.", keep=True),
        SIGN(),
    ]
    return b


CONTRACTS = {
    "HDLD": dict(blocks=hop_dong_lao_dong, ten="Hợp đồng lao động Full-time",
                 ten_file="Hợp_đồng_lao_động_Fulltime", template="TEMPLATE_Hop_dong_lao_dong_Fulltime.docx"),
    "HDKS": dict(blocks=hop_dong_dich_vu, ten="Hợp đồng dịch vụ kỹ sư",
                 ten_file="Hợp_đồng_dịch_vụ_kỹ_sư", template="TEMPLATE_Hop_dong_dich_vu_ky_su.docx"),
}
