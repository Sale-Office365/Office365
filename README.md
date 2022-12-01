# OFFICE 365 #

## A. OFFICE 365 MONDO 2016 ##

- Dùng AIO TOOLS V3.1.3
- Download AIO Tools V3.13 [tại đây](https://1drv.ms/u/s!AkwSBX-xWiVhgReolwU8a9uuJrz7?e=AyNym8), giải nén, rồi run file Activate AIO Tools v3.1.3 by Savio.cmd bằng quyền Run Administrator

![image](https://user-images.githubusercontent.com/103977676/200775134-c098a7a2-a110-4dc0-b268-7b9e24a8d0b3.png)

Bấm phím 5

![image](https://user-images.githubusercontent.com/103977676/200775305-a56f4a53-1bfe-45d3-be11-9f8a06738a3e.png)

Sau khi convert sang Office 365 Mondo 2016 xong, ta download file txt [tại đây](https://1drv.ms/t/s!AkwSBX-xWiVhiSWxpFve3dWSU3CR?e=unmxtg), rồi mở lên bấm Save As với tên kichhoat.cmd run fle này dưới quyền Run Administrator

- Tạo tác vụ gia hạn kích hoạt vĩnh viễn:

![image](https://user-images.githubusercontent.com/103977676/200756492-50b60776-f99b-4e12-8352-090c14850910.png)

Bấm phím O để tạo tác vụ rồi làm theo đề xuất.

## B. OFFICE 365 ENTERPRISE, 365 PROLUS ##

- Kích hoạt bằng tài khoản: Office 365 A1 Plus, A3, A5, E3, E5...
- Các tài khoản tạo sẵn: [bấm vào đây](https://bsthanh-my.sharepoint.com/:w:/g/personal/laptopxiaomi_bsthanh_tk/EQa9vlOr8JdOqcUEYGyjjfQBvW7eHmeqtjR1KMf__A2lHw?e=YgQkSj)
- Tham khảo thêm về office 365 E5 [bấm vào đây](https://github.com/BsNgChiThanh/Tao-office-365-E5-kich-hoat-Office-365-for-desktop), ngày nay cụ Microsoft không cho Việt Nam đăng ký, nếu các bạn muốn hãy nhờ dùng số điện thoại nước ngoài để xác minh.
- Đăng nhập các tài khoản này vào để khích hoạt.
- Done!

# GHI CHÚ: #

Bấm vô biểu tượng Windows trên màn hình desktop chọn Windows System rồi chọn Command Prompt chọn More chọn Run As Administrator

![image](https://user-images.githubusercontent.com/103977676/205012427-c8458db5-51cf-4f10-9acc-637140df63b6.png)

Sau đó tùy Win mà dán câu lệnh:

Windows 32 Bit

```php
cd /d %ProgramFiles(x86)%\Microsoft Office\Office16
```

Windows 64 Bit

```php
cd /d %ProgramFiles%\Microsoft Office\Office16
```

Sau đó Install giấy phép Office 2019 Prolus Volume License:

```php
for /f %x in ('dir /b ..\root\Licenses16\proplusvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%x"
```

Kích hoạt Office 2019 Pro Plus bằng key KMS lần lượt dán câu lệnh:

Câu 1:

```php
cscript ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99
```

Câu 2:

```php
cscript ospp.vbs /unpkey:BTDRB >nul
```

Câu 3:

```php
cscript ospp.vbs /unpkey:KHGM9 >nul
```

Câu 4:

```php
cscript ospp.vbs /unpkey:CPQVG >nul
```

Câu 5:

```php
cscript ospp.vbs /unpkey:WFG99 >nul
```

Câu 6:

```php
cscript ospp.vbs /sethst:s8.uk.to
```

Câu 7:

```php
cscript ospp.vbs /setprt:1688
```

Câu 8:

```php
cscript ospp.vbs /act
```

## Vậy khi thành công thì Office đã kích hoạt 180 ngày! ##
