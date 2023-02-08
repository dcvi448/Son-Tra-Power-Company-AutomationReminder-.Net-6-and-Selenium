A. Cấu hình bắt buộc
1. Thư mục tải xuống dữ liệu:
Vào cài đặt Chrome gõ "download" vào ô tìm kiếm, lấy đường dẫn mặc định mà Chrome sử dụng sau đó mở tệp config.json (trong thư mục chạy app), điền đường dẫn vào dòng folderDir
2. Kiểm tra user và password đăng nhập tại trang ktat đã đúng chưa
Mở tệp config.json (trong thư mục chạy app), tìm cấu trúc này và kiểm tra mật khẩu đã đúng chưa?
    "username": {
        "user": "XXXXXXXX"
    },
    "password": {
        "pass": "XXXXXXXX"
    },
3. Thiết lập token eoffice, cấu hình các thiết bị gửi thông báo, danh sách các user eoffice nhận tin nhắn
Mở tệp config.json (trong thư mục chạy app),
a) Thiết lập token eoffice dựa vào "eoffice": { "token": XXXXX }
b) Thiết lập thông báo ngày đến hạn thiết bị qua eoffice từ ngày nào tới ngày nào & danh sách các user eoffice nhận tin nhắn
   "alert": {
        "fromDay": "2",
        "toDay": "10",
         "users": ["thienhv"],
    }