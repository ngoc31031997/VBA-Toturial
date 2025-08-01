
# 🔤 Hướng Dẫn Đầy Đủ Về Xử Lý Chuỗi Trong VBA

VBA cung cấp nhiều hàm xử lý chuỗi giúp thao tác văn bản hiệu quả.

---

## 🔎 `InStr` – Tìm vị trí chuỗi con

```vba
InStr([start], string1, string2, [compare])
```

- Trả về vị trí xuất hiện đầu tiên của `string2` trong `string1`
- Trả về 0 nếu không tìm thấy

**Ví dụ:**

```vba
InStr("hello world", "world") '→ 7
```

---

## 🔄 `InStrRev` – Tìm từ phải sang trái

```vba
InStrRev(string1, string2, [start], [compare])
```

- Giống `InStr` nhưng tìm từ phải sang trái

**Ví dụ:**

```vba
InStrRev("abc_def_ghi", "_") '→ 8
```

---

## 🔡 `LCase` – Viết thường

```vba
LCase(string)
```

**Ví dụ:**

```vba
LCase("XIN CHÀO") '→ "xin chào"
```

---

## 🔠 `UCase` – Viết hoa

```vba
UCase(string)
```

**Ví dụ:**

```vba
UCase("hello") '→ "HELLO"
```

---

## 🔙 `Left` – Lấy ký tự từ trái

```vba
Left(string, length)
```

**Ví dụ:**

```vba
Left("abcdef", 3) '→ "abc"
```

---

## 🔚 `Right` – Lấy ký tự từ phải

```vba
Right(string, length)
```

**Ví dụ:**

```vba
Right("abcdef", 3) '→ "def"
```

---

## 🔍 `Mid` – Cắt chuỗi ở giữa

```vba
Mid(string, start, [length])
```

**Ví dụ:**

```vba
Mid("abcdef", 2, 3) '→ "bcd"
```

---

## 🧹 `LTrim`, `RTrim`, `Trim` – Xóa khoảng trắng

```vba
LTrim(string)  'Xóa bên trái
RTrim(string)  'Xóa bên phải
Trim(string)   'Xóa hai bên
```

**Ví dụ:**

```vba
Trim("  hello  ") '→ "hello"
```

---

## 🔢 `Len` – Độ dài chuỗi

```vba
Len(string)
```

**Ví dụ:**

```vba
Len("abc") '→ 3
```

---

## 🔁 `Replace` – Thay thế chuỗi con

```vba
Replace(expression, find, replace, [start], [count], [compare])
```

**Ví dụ:**

```vba
Replace("1-2-3", "-", "/") '→ "1/2/3"
```

---

## ⬜ `Space` – Tạo chuỗi khoảng trắng

```vba
Space(n)
```

**Ví dụ:**

```vba
"Hello" & Space(3) & "World" '→ "Hello   World"
```

---

## 🔍 `StrComp` – So sánh chuỗi

```vba
StrComp(string1, string2, [compare])
```

- Trả về `0`: giống nhau, `-1`: nhỏ hơn, `1`: lớn hơn

**Ví dụ:**

```vba
StrComp("a", "A", vbTextCompare) '→ 0
```

---

## 📏 `String` – Lặp ký tự

```vba
String(number, character)
```

**Ví dụ:**

```vba
String(5, "*") '→ "*****"
```

---

## 🔁 Hàm tự viết: Đảo chuỗi

```vba
Function ReverseStr(str As String) As String
    Dim i As Integer, result As String
    For i = Len(str) To 1 Step -1
        result = result & Mid(str, i, 1)
    Next i
    ReverseStr = result
End Function
```

**Ví dụ:**
```vba
ReverseStr("abc") '→ "cba"
```

---

Bạn có muốn tách riêng từng nhóm hàm theo chủ đề (ví dụ: thao tác vị trí, thao tác định dạng, thao tác xóa…) không?
