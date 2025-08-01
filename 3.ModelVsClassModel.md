
# 🧩 So sánh Module và Class Module trong VBA

Module và Class Module có mục đích và cách sử dụng khác nhau trong VBA.

---

## 🛠️ Module (Mô-đun Tiêu chuẩn)

**Ý nghĩa:** Chứa các thủ tục (Sub) và hàm (Function) dùng chung.  
**Đặc điểm:**  
- Code tĩnh, không tạo được đối tượng.  
- Dùng cho macro, hàm tiện ích.

**Ví dụ:**
```vba
Sub ShowMessage()
    MsgBox "Đây là Module tiêu chuẩn!", vbInformation
End Sub
```

---

## 🛠️ Class Module (Mô-đun Lớp)

**Ý nghĩa:** Định nghĩa lớp, hỗ trợ lập trình hướng đối tượng (OOP).  

**Đặc điểm:**  
- Tạo đối tượng với thuộc tính, phương thức, sự kiện.  
- Dùng để mô phỏng thực thể (ví dụ: Nhân viên, Sản phẩm).

**Ví dụ:**

*Class Module: `Employee`*
```vba
Private pName As String

Public Property Let Name(value As String)
    pName = value
End Property

Public Property Get Name() As String
    Name = pName
End Property

Public Sub ShowInfo()
    MsgBox "Tên nhân viên: " & pName
End Sub
```

*Module: `Test`*
```vba
Sub TestClass()
    Dim emp As New Employee
    emp.Name = "Nguyễn Văn A"
    emp.ShowInfo
End Sub
```

---

## 📌 So sánh

| **Tiêu chí**   | **Module**                       | **Class Module**                             |
|----------------|----------------------------------|----------------------------------------------|
| **Loại code**  | Tĩnh, không tạo đối tượng        | Hướng đối tượng, tạo đối tượng               |
| **Ứng dụng**   | Hàm/macro chung                  | Mô phỏng thực thể (Nhân viên, Sản phẩm)      |
| **Ví dụ**      | Hàm tính toán, macro đơn giản    | Quản lý danh sách nhân viên                  |

---

## ⚠️ Lưu ý

- **Module:** Phù hợp cho code đơn giản, dùng chung.  
- **Class Module:** Dùng khi cần mô hình hóa đối tượng phức tạp.
