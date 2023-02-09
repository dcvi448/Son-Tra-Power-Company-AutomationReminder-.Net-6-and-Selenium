A. Cấu hình bắt buộc

1. Thư mục tải xuống dữ liệu:
Vào cài đặt Chrome gõ "download" vào ô tìm kiếm, lấy đường dẫn mặc định mà Chrome sử dụng sau đó mở tệp config.json (trong thư mục chạy app), điền đường dẫn vào dòng folderDir

LƯU Ý: Với các hệ thống Windows nếu đường dẫn có dấu " " vui lòng dùng kí tự / thay vì \, ví dụ: "C:\Users\Thanh Duc" sẽ chuyển thành "C:/Users/Thanh Duc"

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

-----------------------
B. Cấu hình tuỳ chọn

1. Thay đổi nội dung thông báo tới EOffice, nếu cần truyền các tham số dùng các tham số tương ứng với dữ liệu thiết bị như sau
{deviceName}: tên thiết bị
{fromDay} : từ ngày hết hạn
{toDay}: tới ngày hết hạn
{sumOfDevice}: tổng số thiết bị hết hạn
{unitName}: tên đơn vị
{lastestUsedDate}: ngày đưa vào sử dụng
{customerName}: tên khách hàng
{deviceProductionCode}: mã hiệu sản xuất
{deviceManufacturerName}: hãng sản xuất
{deviceManufactureCountry}: nước sản xuất
{deviceYearOf}: năm sản xuất
{deviceStatus}: tình trạng thiết bị

2. Lọc trạng thái thiết bị
Dựa vào 
"status": {
        "choose": "4" //4 là gần hết hạn, 2 là hết hạn, -1 là hỏng, 3 là chưa xác minh 
},

3. Cấu hình kết quả có bao gồm các đơn vị trực thuộc hay không?
Dựa vào
"unit": {
        "includeAll": "yes",
// yes có nghĩa là có, no nghĩa là không

4. Thiết lập tuỳ chọn lọc thiết bị theo đơn vị, loại, tổ đội, cấp điện áp, trạng thái...
Dựa vào
 "unitName": đơn vị
 "deviceType": loại thiết bị
 "device": thiết bị
 "levelA": cấp điện áp
 "team": tổ đội
 "toDate": đến ngày
 
 Với mỗi tuỳ chọn trên đều có thiết lập

 "choose": chọn đối tượng nào trong danh sách bằng mã của nó (liên hệ để biết cách)
 "isUse": "no" nghĩa là bỏ qua (để mặc định), "yes" là thực hiện chọn dựa vào mã của "choose" ở trên


5. Thay đổi id đối tượng (NÂNG CAO), sử dụng trong trường hợp cấu trúc trang ktat thay đổi
Trong tất cả các id quan trọng luôn được xử lý flexible, tức là dựa theo id đã được cấu hình, nếu id ở trang ktat thay đổi, bạn chỉ cần sửa lại id ở đây
chứ không cần phải sửa chương trình.
Ví dụ:
"username": { 
    "id": "Txtusername" ...
"password": { 
    "id": "Txtpass", ...
"login": {
    "id": "Tcap" ...
"status": {
    "id": "ctl00_ContentPlaceHolder1_drCTT" ...
"getAllDevice": {
    "id": "ctl00_ContentPlaceHolder1_btnLayds" ...
"exportXLSX": {
    "id": "grdDsPhieuCT_DXCTMenu0_DXI1_T" ...

