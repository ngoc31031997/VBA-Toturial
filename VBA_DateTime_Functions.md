
# 📅 Hướng Dẫn Đầy Đủ Về Hàm Ngày Giờ Trong VBA

VBA hỗ trợ nhiều hàm để xử lý ngày giờ – từ kiểm tra, tính toán đến định dạng.

---

## 📆 NHÓM HÀM NGÀY (DATE FUNCTIONS)

### 🔹 `Date`

- Trả về ngày hiện tại trên hệ thống.

```vba
MsgBox Date '→ 01/08/2025
```

---

### 🔹 `CDate`

- Chuyển chuỗi hoặc giá trị về kiểu ngày.

```vba
CDate("01/08/2025") '→ #01/08/2025#
```

---

### 🔹 `DateAdd`

- Cộng thêm thời gian vào một ngày.

```vba
DateAdd(interval, number, date)
```

**Ví dụ:**

```vba
DateAdd("m", 1, "01/08/2025") '→ 01/09/2025
```

---

### 🔹 `DateDiff`

- Tính khoảng cách giữa 2 ngày theo đơn vị.

```vba
DateDiff(interval, date1, date2)
```

**Ví dụ:**

```vba
DateDiff("d", "01/01/2025", "10/01/2025") '→ 9
```

---

### 🔹 `DatePart`

- Trích xuất thành phần từ ngày (như quý, tuần, tháng...).

```vba
DatePart("q", "15/08/2025") '→ 3 (Quý 3)
```

---

### 🔹 `DateSerial`

- Tạo ngày từ các thành phần số.

```vba
DateSerial(2025, 8, 1) '→ #01/08/2025#
```

---

### 🔹 `Format` / `FormatDateTime`

- Định dạng ngày giờ theo mẫu.

```vba
Format(Date, "dd-mm-yyyy") '→ "01-08-2025"
FormatDateTime(Now, vbLongDate) '→ "Friday, August 1, 2025"
```

---

### 🔹 `IsDate`

- Kiểm tra xem giá trị có phải ngày hợp lệ không.

```vba
IsDate("01/08/2025") '→ True
IsDate("abc") '→ False
```

---

### 🔹 `Day`, `Month`, `Year`

```vba
Day(#01/08/2025#)   '→ 1
Month(#01/08/2025#) '→ 8
Year(#01/08/2025#)  '→ 2025
```

---

### 🔹 `WeekDay`, `WeekDayName`, `MonthName`

```vba
WeekDay(#01/08/2025#)       '→ 6 (Thứ Sáu)
WeekDayName(6)              '→ "Friday"
MonthName(8)                '→ "August"
```

---

## ⏰ NHÓM HÀM GIỜ (TIME FUNCTIONS)

### 🔹 `Now`

- Trả về ngày giờ hiện tại.

```vba
MsgBox Now '→ 01/08/2025 10:00:00
```

---

### 🔹 `Hour`, `Minute`, `Second`

```vba
Hour(Now)   '→ 10
Minute(Now) '→ 0
Second(Now) '→ 0
```

---

### 🔹 `Time`

- Trả về giờ hiện tại.

```vba
MsgBox Time '→ 10:00:00
```

---

### 🔹 `Timer`

- Trả về số giây từ 0:00 đêm đến thời điểm hiện tại.

```vba
Debug.Print Timer '→ 36000 (10 giờ sáng)
```

---

### 🔹 `TimeSerial`

- Tạo giá trị thời gian từ giờ, phút, giây.

```vba
TimeSerial(10, 30, 0) '→ 10:30:00 AM
```

---

### 🔹 `TimeValue`

- Chuyển chuỗi thành giờ.

```vba
TimeValue("14:45:00") '→ 2:45:00 PM
```

---

Bạn có muốn chia các nhóm thành file riêng: `DateFunctions.md`, `TimeFunctions.md`?
