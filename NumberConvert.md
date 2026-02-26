## NumberConvert.bas — Hướng dẫn sử dụng

### Giới thiệu
`NumberConvert.bas` là module VBA cung cấp 2 hàm chuyển đổi số thành chữ:

- Chuyển số thành chữ tiếng Anh  
- Chuyển số thành chữ tiếng Việt  

Thường dùng để hiển thị số tiền bằng chữ trong chứng từ và báo cáo kế toán.

## Import module vào Excel

1. Mở Excel  
2. Nhấn **Alt + F11** để mở VBA Editor  
3. Chọn **File → Import File…**  
4. Chọn file `NumberConvert.bas`  

Sau khi import, module `NumberConvert` sẽ xuất hiện trong Project.

## Cách sử dụng trong Excel

### Đọc số thành chữ tiếng Anh

```excel
=NumberToWordsEnglish(A1)
```

### Ví dụ

| Giá trị | Kết quả |
|---|---|
| 1234 | One thousand two hundred thirty four |

---

### Đọc số thành chữ tiếng Việt

```excel
=NumberToWordVietnamese(A1)
```

### Ví dụ

| Giá trị | Kết quả |
|---|---|
| 1234 | Một nghìn hai trăm ba mươi bốn |

---

## Ứng dụng

- In số tiền bằng chữ trên hóa đơn  
- Phiếu thu / phiếu chi  
- Báo cáo tài chính  
- Tài liệu kế toán
