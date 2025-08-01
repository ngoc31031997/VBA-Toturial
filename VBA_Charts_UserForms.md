
# 📊 Hướng Dẫn VBA – Vẽ Biểu Đồ & Thiết Kế UserForm

---

## 📈 PHẦN 1 – PROGRAMMING CHARTS (Biểu đồ trong VBA)

### 🔹 Thêm biểu đồ mới

```vba
Dim chtObj As ChartObject
Set chtObj = ActiveSheet.ChartObjects.Add(Left:=100, Width:=400, Top:=50, Height:=300)
chtObj.Chart.ChartType = xlColumnClustered
chtObj.Chart.SetSourceData Source:=Range("A1:B5")
```

---

### 🔹 Định dạng biểu đồ

```vba
With chtObj.Chart
    .HasTitle = True
    .ChartTitle.Text = "Doanh thu theo tháng"
    .Axes(xlCategory).HasTitle = True
    .Axes(xlCategory).AxisTitle.Text = "Tháng"
    .Axes(xlValue).AxisTitle.Text = "Doanh thu"
End With
```

---

### 🔹 Các loại biểu đồ phổ biến

| Loại                     | Mã VBA               |
|--------------------------|----------------------|
| Cột                      | `xlColumnClustered`  |
| Đường                   | `xlLine`             |
| Tròn                     | `xlPie`              |
| Thanh ngang             | `xlBarClustered`     |
| Kết hợp (phức tạp hơn)   | Kết hợp nhiều series |

---

### 🔹 Xóa biểu đồ

```vba
chtObj.Delete
```

📌 Biểu đồ là đối tượng nằm trên worksheet (`ChartObject`) hoặc riêng (`Chart` sheet).

---

## 🧩 PHẦN 2 – USER FORMS (Giao diện người dùng tùy chỉnh)

UserForm giúp tạo giao diện tương tác – nhập liệu, chọn tùy chọn, xác nhận hành động.

---

### 🛠️ Tạo UserForm

- Vào **VBA Editor (Alt+F11)** → Insert → **UserForm**
- Thêm các điều khiển (TextBox, Label, Button, ComboBox...)

---

### 🔹 Code xử lý trong UserForm

```vba
Private Sub cmdOK_Click()
    MsgBox "Xin chào, " & txtName.Value
End Sub
```

---

### 🔹 Gọi UserForm từ Module

```vba
Sub ShowForm()
    UserForm1.Show
End Sub
```

---

### 🔹 Truy cập dữ liệu từ UserForm

```vba
Dim ten As String
ten = UserForm1.txtName.Value
```

Hoặc gán dữ liệu vào Form trước khi hiển thị:

```vba
UserForm1.txtName.Value = "Ngọc"
UserForm1.Show
```

---

### 📦 Các control phổ biến

| Control     | Công dụng               |
|-------------|-------------------------|
| Label       | Hiển thị văn bản        |
| TextBox     | Nhập liệu               |
| CommandButton | Nút bấm                |
| ComboBox    | Danh sách chọn          |
| ListBox     | Danh sách nhiều lựa chọn |
| Frame       | Nhóm điều khiển         |

---

### 🧯 Đóng Form

```vba
Unload Me 'trong chính form
Unload UserForm1 'từ ngoài
```

---

## 📌 Lưu ý

- Có thể tùy chỉnh màu sắc, font, kích thước control trong cửa sổ Properties
- Sử dụng `.Hide` thay vì `Unload` nếu muốn giữ giá trị đã nhập
- Dùng `Initialize` để đặt giá trị mặc định khi Form mở

---

Bạn cần ví dụ cụ thể cho UserForm nhập dữ liệu và ghi vào sheet không?
