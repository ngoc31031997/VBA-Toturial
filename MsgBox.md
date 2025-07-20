Hướng Dẫn MsgBox trong VBA Excel
Hàm MsgBox trong VBA Excel hiển thị hộp thoại thông báo, cho phép tương tác với người dùng thông qua nút, biểu tượng, và tiêu đề tùy chỉnh.
Cú pháp
MsgBox(Prompt, [Buttons], [Title], [HelpFile], [Context])

Các tham số

Prompt (Bắt buộc):

Chuỗi văn bản hiển thị trong hộp thoại.
Độ dài tối đa: ~1024 ký tự.
Ví dụ: "Đây là thông báo!".


Buttons (Tùy chọn):

Quy định kiểu nút, biểu tượng, nút mặc định, và chế độ hiển thị.
Kiểu nút:
vbOKOnly (0): Chỉ nút OK.
vbOKCancel (1): Nút OK và Cancel.
vbAbortRetryIgnore (2): Nút Abort, Retry, Ignore.
vbYesNoCancel (3): Nút Yes, No, Cancel.
vbYesNo (4): Nút Yes và No.
vbRetryCancel (5): Nút Retry và Cancel.


Biểu tượng:
vbCritical (16): Dấu X (lỗi).
vbQuestion (32): Dấu hỏi.
vbExclamation (48): Dấu chấm than (cảnh báo).
vbInformation (64): Chữ "i" (thông tin).


Nút mặc định:
vbDefaultButton1 (0): Nút đầu tiên mặc định.
vbDefaultButton2 (256): Nút thứ hai mặc định.
vbDefaultButton3 (512): Nút thứ ba mặc định.


Chế độ hiển thị:
vbApplicationModal (0): Khóa Excel cho đến khi trả lời.
vbSystemModal (4096): Khóa toàn bộ hệ thống.


Căn chỉnh văn bản (ít dùng):
vbMsgBoxRight (524288): Căn phải.
vbMsgBoxRtlReading (1048576): Đọc từ phải sang trái.


Ví dụ: vbYesNo + vbQuestion hiển thị nút Yes/No và biểu tượng dấu hỏi.


Title (Tùy chọn):

Văn bản trên thanh tiêu đề.
Mặc định: Tên ứng dụng (ví dụ: "Microsoft Excel").
Ví dụ: "Xác nhận hành động".


HelpFile (Tùy chọn):

Đường dẫn đến tệp trợ giúp (.chm).
Thường để trống.


Context (Tùy chọn):

Số định danh chủ đề trợ giúp trong HelpFile.
Chỉ dùng khi có HelpFile.



Giá trị trả về

vbOK (1): Nhấn OK.
vbCancel (2): Nhấn Cancel.
vbAbort (3): Nhấn Abort.
vbRetry (4): Nhấn Retry.
vbIgnore (5): Nhấn Ignore.
vbYes (6): Nhấn Yes.
vbNo (7): Nhấn No.

Ví dụ
Sub TestMsgBox()
    Dim response As VbMsgBoxResult
    response = MsgBox("Bạn có muốn lưu file?", vbYesNo + vbInformation, "Lưu File")
    If response = vbYes Then
        MsgBox "File đã được lưu!", vbOKOnly, "Thông báo"
    Else
        MsgBox "Hủy lưu file!", vbOKOnly, "Thông báo"
    End If
End Sub

Lưu ý

Kết hợp Buttons bằng dấu +.
Dùng vbApplicationModal để khóa Excel.
Tránh vbSystemModal vì khóa toàn hệ thống.
HelpFile và Context hiếm dùng.
