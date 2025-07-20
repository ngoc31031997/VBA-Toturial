So sánh Module và Class Module trong VBA
1. Module (Mô-đun Tiêu chuẩn)

Mục đích: Chứa các thủ tục (Sub) và hàm (Function) dùng chung.
Đặc điểm:
Code tĩnh, không tạo được đối tượng.
Dùng cho macro, hàm tiện ích.


Ví dụ:Sub ShowMessage()
    MsgBox "Đây là Module tiêu chuẩn!", vbInformation
End Sub



2. Class Module ( Մô-đun Lớp)

Mục đích: Định nghĩa lớp, hỗ trợ lập trình hướng đối tượng (OOP).

Đặc điểm:

Tạo đối tượng với thuộc tính (Properties), phương thức (Methods), và sự kiện (Events).
Dùng để mô phỏng các thực thể (ví dụ: Nhân viên, Sản phẩm).


Ví dụ:
' Class Module: Employee
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

' Module: Test
Sub TestClass()
    Dim emp As New Employee
    emp.Name = "Nguyễn Văn A"
    emp.ShowInfo
End Sub



So sánh

Module: Code tĩnh, không tạo đối tượng, dùng cho hàm/macro chung.
Class Module: Code hướng đối tượng, tạo đối tượng, dùng cho mô hình hóa thực thể.
Ứng dụng:
Module: Hàm tính toán, macro đơn giản.
Class Module: Quản lý đối tượng phức tạp (ví dụ: danh sách nhân viên).


