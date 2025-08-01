
# 📘 Biến (Variable) và Hằng số (Constant) trong VBA

Hiểu và sử dụng đúng **biến** và **hằng số** là nền tảng khi lập trình VBA Excel.

---

## 🔹 Biến (`Variable`)

Biến là nơi lưu trữ giá trị có thể thay đổi trong quá trình chạy chương trình.

### 🛠️ Khai báo biến

```vba
Dim ten_bien As KieuDuLieu
```

**Ví dụ:**

```vba
Dim tong As Integer
Dim hoTen As String
Dim diemTB As Double
```

### 🧪 Một số kiểu dữ liệu phổ biến

| Kiểu dữ liệu | Mô tả                    | Ví dụ           |
|--------------|---------------------------|------------------|
| `Integer`    | Số nguyên (-32k đến 32k)  | `Dim a As Integer` |
| `Long`       | Số nguyên lớn hơn         | `Dim id As Long` |
| `Double`     | Số thực có dấu            | `Dim x As Double` |
| `String`     | Chuỗi ký tự               | `Dim s As String` |
| `Boolean`    | Đúng/Sai (`True/False`)   | `Dim flag As Boolean` |
| `Date`       | Ngày/giờ                  | `Dim ngay As Date` |
| `Variant`    | Tự động xác định kiểu     | `Dim a` (không cần kiểu) |

### 💡 Cách gán giá trị

```vba
ten_bien = GiaTri
```

**Ví dụ:**

```vba
tong = 10
hoTen = "Nguyễn Văn A"
flag = True
```

---

## 🔹 Hằng số (`Constant`)

Hằng số là giá trị cố định, không thay đổi trong suốt chương trình.

### 🛠️ Khai báo hằng số

```vba
Const ten_hang As KieuDuLieu = GiaTri
```

**Ví dụ:**

```vba
Const PI As Double = 3.14159
Const APP_NAME As String = "Quản lý chi phí"
```

### ⚠️ Lưu ý

- Tên hằng thường viết IN HOA để dễ phân biệt: `MAX_SCORE`, `TAX_RATE`
- Không thể gán lại giá trị cho `Const` sau khi khai báo

---

## 🔍 Ví dụ tổng hợp

```vba
Sub DemoVariableConstant()
    Const MAX_SCORE As Integer = 100
    Dim ten As String
    Dim diem As Integer

    ten = "Lan"
    diem = 85

    MsgBox ten & " đạt " & diem & "/" & MAX_SCORE, vbInformation, "Kết quả"
End Sub
```

---

## ✅ Gợi ý

- Sử dụng `Option Explicit` đầu module để bắt buộc khai báo biến → tránh lỗi chính tả
- Luôn đặt tên biến/hằng có ý nghĩa rõ ràng
- Dùng `Variant` chỉ khi thật sự cần thiết

---

## 🔸 Static, Public, Private

### 🧷 `Static`

Biến `Static` giữ nguyên giá trị giữa các lần chạy `Sub` hoặc `Function`.

```vba
Sub DemSoLanChay()
    Static count As Integer
    count = count + 1
    MsgBox "Số lần chạy: " & count
End Sub
```

> 🔁 Mỗi lần gọi lại `DemSoLanChay`, biến `count` vẫn giữ giá trị cũ.

---

### 🌐 `Public`

- Biến hoặc hằng được khai báo toàn cục, dùng được ở mọi module.
- Thường khai báo trong **Module chuẩn**, ngoài bất kỳ `Sub`/`Function` nào.

```vba
Public userName As String
Public Const VERSION As String = "1.0"
```

---

### 🔒 `Private`

- Biến/hằng chỉ dùng được trong module khai báo.
- Giúp bảo vệ dữ liệu và tránh xung đột tên biến.

```vba
Private dbPassword As String
Private Const TAX_RATE As Double = 0.1
```

---

### 📌 So sánh nhanh

| Phạm vi     | Từ khóa     | Sử dụng ở đâu                     |
|-------------|--------------|-----------------------------------|
| Toàn cục     | `Public`     | Trong Module chuẩn                |
| Cục bộ       | `Private`    | Trong Module/Form/Class           |
| Giữ giá trị | `Static`     | Trong `Sub`/`Function` nội bộ     |

