
# ❗ Hướng Dẫn VBA – Xử Lý Lỗi (Error Handling)

Lỗi trong VBA có thể là cú pháp, thời gian chạy hoặc logic. Hiểu cách xử lý giúp chương trình ổn định và chuyên nghiệp hơn.

---

## 🚫 Syntax Errors (Lỗi cú pháp)

- Xảy ra khi viết sai cú pháp VBA.
- VBA sẽ thông báo ngay khi viết code (chữ đỏ, không chạy được).

**Ví dụ:**

```vba
If x = 5 ' Thiếu Then → lỗi cú pháp
```

---

## ⚠️ Runtime Errors (Lỗi thời gian chạy)

- Xảy ra khi chương trình đang thực thi, ví dụ chia cho 0, truy cập mảng sai chỉ số.

**Ví dụ:**
```vba
Dim a As Integer
a = 5 / 0 ' → lỗi chia cho 0
```

---

## ❓ Logical Errors (Lỗi logic)

- Chương trình **chạy được** nhưng cho kết quả **sai**.

**Ví dụ:**
```vba
If diem >= 5 Then
    MsgBox "Rớt" '→ Sai logic, điều kiện đúng nhưng kết luận sai
End If
```

---

## 🧱 Đối tượng `Err`

VBA cung cấp đối tượng `Err` để truy cập thông tin lỗi.

```vba
Err.Number   '→ Mã lỗi
Err.Description '→ Mô tả lỗi
```

---

## 🧯 Xử lý lỗi – Error Handling

### 🔹 Cấu trúc cơ bản

```vba
On Error GoTo XuLyLoi

' --- Code chính --
Dim a As Integer
a = 5 / 0

Exit Sub

XuLyLoi:
MsgBox "Có lỗi: " & Err.Description
```

### 🔹 `On Error Resume Next`

- Bỏ qua lỗi, tiếp tục dòng tiếp theo.

```vba
On Error Resume Next
a = 5 / 0 ' Không báo lỗi, nhưng a = 0
```

📌 Dùng cẩn thận – có thể bỏ qua lỗi nghiêm trọng mà không biết.

### 🔹 `On Error GoTo 0`

- Tắt xử lý lỗi tùy chỉnh, quay về mặc định.

```vba
On Error GoTo 0
```

---

## 📌 Lưu ý & Mẹo

- Luôn dùng `Exit Sub/Function` **trước nhãn lỗi** để không chạy vào phần xử lý lỗi khi không có lỗi.
- Có thể ghi log lỗi vào file hoặc sheet để dễ tra cứu.
- Ưu tiên xử lý rõ ràng với `GoTo`, tránh lạm dụng `Resume Next`.

---

Bạn có muốn bổ sung ví dụ nâng cao như xử lý lỗi vòng lặp hoặc kết hợp ghi log không?
