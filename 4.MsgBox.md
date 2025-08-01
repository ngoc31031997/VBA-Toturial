# 🧩 Hướng Dẫn MsgBox trong VBA Excel

Hàm `MsgBox` hiển thị hộp thoại tương tác với người dùng (nút, biểu tượng, tiêu đề…).

---

## 🛠️ Cú pháp

```vba
MsgBox(Prompt, [Buttons], [Title], [HelpFile], [Context])
```

---

## 📌 Tham số

### 🔹 `Prompt` *(bắt buộc)*

- Chuỗi hiển thị trong hộp thoại.  
- Tối đa ~1024 ký tự.  
- **Ví dụ:** `"Đây là thông báo!"`

---

### 🔹 `Buttons` *(tùy chọn)*

**1. Kiểu nút:**

| Tên hằng số           | Giá trị | Mô tả                     |
|-----------------------|---------|---------------------------|
| `vbOKOnly`            | 0       | Chỉ nút OK                |
| `vbOKCancel`          | 1       | OK & Cancel               |
| `vbAbortRetryIgnore`  | 2       | Abort, Retry, Ignore      |
| `vbYesNoCancel`       | 3       | Yes, No, Cancel           |
| `vbYesNo`             | 4       | Yes & No                  |
| `vbRetryCancel`       | 5       | Retry & Cancel            |

**2. Biểu tượng:**

| Tên hằng số       | Giá trị | Biểu tượng     |
|-------------------|---------|----------------|
| `vbCritical`      | 16      | ❌ Lỗi          |
| `vbQuestion`      | 32      | ❓ Hỏi           |
| `vbExclamation`   | 48      | ⚠️ Cảnh báo     |
| `vbInformation`   | 64      | ℹ️ Thông tin     |

**3. Nút mặc định:**

| Tên hằng số         | Giá trị | Mô tả                    |
|---------------------|---------|--------------------------|
| `vbDefaultButton1`  | 0       | Mặc định nút 1           |
| `vbDefaultButton2`  | 256     | Mặc định nút 2           |
| `vbDefaultButton3`  | 512     | Mặc định nút 3           |

**4. Chế độ hiển thị:**

| Tên hằng số         | Giá trị | Mô tả                     |
|---------------------|---------|---------------------------|
| `vbApplicationModal`| 0       | Khóa Excel tới khi trả lời |
| `vbSystemModal`     | 4096    | Khóa toàn hệ thống ⚠️      |

**5. Căn chỉnh (hiếm dùng):**

| Tên hằng số           | Giá trị |
|-----------------------|---------|
| `vbMsgBoxRight`       | 524288  |
| `vbMsgBoxRtlReading`  | 1048576 |

🔹 **Ví dụ:**  
```vba
vbYesNo + vbQuestion '→ Hiển thị Yes/No với biểu tượng dấu hỏi
```

---

### 🔹 `Title` *(tùy chọn)*

- Tiêu đề hiển thị trên hộp thoại  
- Mặc định: "Microsoft Excel"  
- **Ví dụ:** `"Xác nhận hành động"`

---

### 🔹 `HelpFile` & `Context` *(hiếm dùng)*

- `HelpFile`: Đường dẫn tệp trợ giúp (.chm)  
- `Context`: ID chủ đề trong tệp trợ giúp

---

## 📤 Giá trị trả về

| Kết quả      | Giá trị |
|--------------|---------|
| `vbOK`       | 1       |
| `vbCancel`   | 2       |
| `vbAbort`    | 3       |
| `vbRetry`    | 4       |
| `vbIgnore`   | 5       |
| `vbYes`      | 6       |
| `vbNo`       | 7       |

---

## 🔍 Ví dụ sử dụng

```vba
Sub TestMsgBox()
    Dim response As VbMsgBoxResult
    response = MsgBox("Bạn có muốn lưu file?", vbYesNo + vbInformation, "Lưu File")

    If response = vbYes Then
        MsgBox "File đã được lưu!", vbOKOnly, "Thông báo"
    Else
        MsgBox "Hủy lưu file!", vbOKOnly, "Thông báo"
    End If
End Sub
```

---

## ⚠️ Lưu ý

- Dùng **`+`** để kết hợp nhiều tùy chọn trong `Buttons`
- Nên dùng `vbApplicationModal` để khóa Excel
- Tránh `vbSystemModal` (khóa cả hệ thống!)
- `HelpFile` và `Context` hầu như không cần thiết
