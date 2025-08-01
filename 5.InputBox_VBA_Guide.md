
# ✏️ Hướng Dẫn InputBox trong VBA Excel

Hàm `InputBox` hiển thị hộp thoại để người dùng nhập dữ liệu (chuỗi, số...).

---

## 🛠️ Cú pháp

```vba
InputBox(Prompt, [Title], [Default], [XPos], [YPos], [HelpFile], [Context])
```

---

## 📌 Tham số

### 🔹 `Prompt` *(bắt buộc)*

- Chuỗi hiển thị trong hộp thoại.  
- Tối đa ~1024 ký tự.  
- **Ví dụ:** `"Nhập tên của bạn:"`

---

### 🔹 `Title` *(tùy chọn)*

- Tiêu đề hiển thị trên hộp thoại.  
- Mặc định: `"Microsoft Excel"`  
- **Ví dụ:** `"Thông tin người dùng"`

---

### 🔹 `Default` *(tùy chọn)*

- Giá trị hiển thị sẵn trong ô nhập.  
- **Ví dụ:** `"Nguyễn Văn A"`  
- Nếu không có, ô sẽ để trống.

---

### 🔹 `XPos`, `YPos` *(tùy chọn)*

- Tọa độ hiển thị hộp thoại tính từ góc trên trái màn hình (pixel).  
- Mặc định: hệ thống tự căn giữa.  
- **Ví dụ:** `XPos:=500, YPos:=300`

---

### 🔹 `HelpFile` & `Context` *(hiếm dùng)*

- `HelpFile`: Đường dẫn tệp trợ giúp (.chm)  
- `Context`: ID chủ đề trong tệp trợ giúp

---

## 📤 Giá trị trả về

- Giá trị người dùng nhập (kiểu `String`)  
- Nếu nhấn Cancel → trả về chuỗi rỗng `""`

---

## 🔍 Ví dụ sử dụng

```vba
Sub TestInputBox()
    Dim userName As String
    userName = InputBox("Nhập tên của bạn:", "Thông tin", "Nguyễn Văn A")

    If userName <> "" Then
        MsgBox "Chào bạn: " & userName, vbInformation, "Kết quả"
    Else
        MsgBox "Bạn đã hủy nhập!", vbExclamation, "Thông báo"
    End If
End Sub
```

---

## ⚠️ Lưu ý

- `InputBox` luôn trả về **chuỗi** (`String`)  
  → Cần kiểm tra kiểu dữ liệu nếu nhập số  
- Để nhập giá trị phức tạp hơn, dùng `Application.InputBox` (nâng cao hơn)  
- Không thể giới hạn độ dài nhập bằng tham số – cần xử lý thủ công sau khi nhận giá trị  
