
# 🧠 Hướng Dẫn VBA – Hàm Tự Định Nghĩa, Sub Procedure và Events

---

## 🔧 PHẦN 1 – USER-DEFINED FUNCTIONS (Hàm tự định nghĩa)

### 🛠️ Định nghĩa Function

```vba
Function TenHam(thamSo1 As Kieu, ...) As KieuTraVe
    ' Xử lý
    TenHam = gia_tri
End Function
```

**Ví dụ:**

```vba
Function Tong(a As Double, b As Double) As Double
    Tong = a + b
End Function
```

---

### 📞 Gọi hàm (Calling a Function)

- Trong Sub hoặc ô Excel:

```vba
Dim kq As Double
kq = Tong(5, 3) '→ 8
```

- Dùng trực tiếp trong Excel:

```excel
=Tong(5, 3)
```

📌 Chỉ có thể dùng trong Excel nếu lưu ở **Module chuẩn (không phải trong Sheet hoặc Form)**

---

## 🧩 PHẦN 2 – SUB PROCEDURE (Thủ tục con)

### 🔧 Định nghĩa Sub

```vba
Sub TenThuTuc()
    ' Các lệnh xử lý
End Sub
```

**Ví dụ:**
```vba
Sub Hello()
    MsgBox "Xin chào VBA!"
End Sub
```

---

### 🔗 Gọi Sub (Calling Procedures)

- Từ Sub khác:

```vba
Call Hello()
```

- Từ sự kiện (event), button, hoặc trực tiếp trong Excel (Alt+F8)

📌 `Sub` không trả giá trị – dùng để thực thi hành động.

---

## ⚡ PHẦN 3 – EVENTS (Sự kiện)

### 📄 Worksheet Events

Các sự kiện liên quan đến 1 sheet cụ thể – khai báo trong cửa sổ `Sheet1`:

```vba
Private Sub Worksheet_Change(ByVal Target As Range)
    MsgBox "Ô vừa thay đổi là: " & Target.Address
End Sub
```

Một số sự kiện thông dụng:

| Sự kiện               | Mô tả                                 |
|------------------------|----------------------------------------|
| `Worksheet_Change`     | Khi ô bị thay đổi                     |
| `Worksheet_SelectionChange` | Khi chọn ô khác                   |
| `Worksheet_Activate`   | Khi mở hoặc kích hoạt Sheet           |

---

### 📁 Workbook Events

Khai báo trong **ThisWorkbook**.

```vba
Private Sub Workbook_Open()
    MsgBox "Chào mừng bạn mở file!"
End Sub
```

Các sự kiện hay dùng:

| Sự kiện             | Mô tả                                 |
|----------------------|----------------------------------------|
| `Workbook_Open`      | Khi file được mở                      |
| `Workbook_BeforeClose` | Trước khi file đóng lại              |
| `Workbook_SheetChange` | Khi có sheet bất kỳ thay đổi         |

---

Bạn muốn tách thành 3 file riêng biệt (Function, Sub, Events) để dễ dùng lại không?
