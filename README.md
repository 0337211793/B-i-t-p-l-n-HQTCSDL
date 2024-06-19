Mục Lục:

CHƯƠNG I. KHẢ NĂNG CỦA CƠ SỞ DỮ LIỆU	
1.Lời nói đầu	
2. Khái niệm về cơ sở dữ liệu và Hệ quản trị cơ sở dữ liệu	
3. Giới thiệu về Access	
4. Mục đích của Quản ly đoàn viên?	

CHƯƠNG II. PHÂN TÍCH CƠ SỞ DỮ LIỆU	
Lược đồ  quan hệ (trang sau):	

CHƯƠNG III. CÁC ĐỐI TƯỢNG	
I. Đối tượng bảng	
1.Các bảng	
2. Các quan hệ	
II . Đối tượng Form	
Hình 32 – form F_Thay mat khau	
III . Macro và hệ thống menu	
1. Macros	
2 . Hệ thống menu	
a. Menu Nhap du lieu	
b.Menu Tien ich	34
c.Menu Danh sach	
d.Menu Timkiem	
e.Menu Bao cao	
f.Menu In an	
g. Menu Thoat
IV .Đối tượng  Querrry	
Hệ thống Querry	
V .Đối tượng  report		
Sử dụng chương trình - Ví dụ	
1. Nhập các thông số ban đầu	
2. Nhập số liệu	
3. Xem dữ liệu	
4.Tìm kiếm	
5.  Báo cáo và thống kê	
6. In ấn	



CHƯƠNG I. KHẢ NĂNG CỦA CƠ SỞ DỮ LIỆU

1.Lời nói đầu
Kỷ nguyên công nghệ thông tin – tên mà người ta vẫn thường gọi kỷ nguyên mà chúng ta đang sống. Khi đi bất cứ đâu ta cũng bắt gặp công nghệ thông tin. Nó xâm nhập vào mọi lĩnh vực của đời sống xã hội, mọi nghành, mọi nghề, mọi lĩnh vực. Ngày nay, không ai còn nghi ngờ gì về vai trò của thông tin trong đời sống, trong kỹ thuật, trong kinh doanh, trong quản lý,…Việc nắm bắt thông tin nhanh, chính xác kịp thời ngày càng có vai trò quyết định cho công tác quản lý điều hành. Nói cách khác ngày nay, Quản Lý thực chất là Quản lý thông tin. Để quản lý được các thông tin đó trên máy tình thì các thông tin đó bao giờ cũng được thể hiện bằng Dữ liệu ( Data) ghi trên dạng tải nào đó. Bởi vậy khi ta nói quản lý thông tin là Quản lý dữ liệu.
2. Khái niệm về cơ sở dữ liệu và Hệ quản trị cơ sở dữ liệu

Trong những năm gần đây thuật ngữ Cơ sở dữ liệu ( Database) không còn xa lạ với những người làmn tin học. Ngày nay các ứng dụng tin học có trong mọi lĩnh vực và ngày càng có xu hướng tăng nhanh. Xu hướng tích cực đó kéo theo ngày càng có đông đảo người tham gia, quan tâm đến thiết kế xây dựng các cơ sở dữ liệu. Ngày nay có rất nhiều mô hình Cơ sở dữ liệu, mỗi mô hình đều có ưu nhược điểm của mình, tuy nhiên mô hình Cơ sở dữ liệu quan hệ (Relational data Model) do E.Codd đề xuất tỏ ra có nhiều ưu điểm khi thiết kế các ứng dụng vừa và nhỏ . 
Dựa trên các mô hình Cơ sở dữ liệu, các hãng máy tính lớn đã xây dựng các Hệ quản trị Cơ sở dữ liệu có nhiều khả năng rất mạnh. Đó là những công cụ tốt cho người lập trình, để giúp họ xây dựng nên các ứng dụng quản lý đa dạng, phục vụ cho mọi yêu cầu của công tác quản lý và điều hành.

Các khái niệm :
a)	Cơ sở dữ liệu ( Database)
Cỏ sỏ dữ liệu là một bộ sưu tập các dữ liệu tác nghiệp được lưu trữ lại và được các hệ hệ ứng dụng của 1 xí nghiệp cụ thể nào đó sử dụng.
Xí nghiệp ( Enterprise) là một thuật ngữ chung để chỉ các hoạt động thương mại, hoạt động khoa học kỹ thuật, hoạt động dịch vụ và các hoạt động khác với quy mô đủ lớn.
b)	Hệ quản trị Cơ sở dữ liệu (Database Management System)
Hệ quản trị cơ sở dữ liệu là một bộ các chương trình cho phép tạo lập và xử lý 
Hệ quản trị cơ sở dữ liệu được coi như một bộ diễn dịch với ngôn ngữ bậc cao nhằm giúp người sử dụng có thể sử dụng được hệ thống mà ít nhiều không cần quan tâm đến thông tin chi tiết hoặc biểu diễn dữ liệu trong máy
Đặc trưng :
-	Đảm bảo tính độc lập dữ liệu tức là tính bất biến của hệ ứng dụng trong chương trình lưu trữ và chiến  lược truy nhập.
-	Có đầy đủ các phép thao tác dữ liệu ( bổ sung, sửa đổi, xoá, tìm kiếm )
-	Có ngôn ngữ giao diện
-	
3. Giới thiệu về Access
Access là một hệ quản trị cơ sở dữ liệu trên môi trường Windows. Nó dựng để tạo ra các cơ sở dữ liệu (chương trình+dữ liệu) cho hầu hết các bài toán thường gặp trong quản lý, thống kê, kế toán, dự toán... bằng các công cụ hữu hiệu của riêng nó.
Microsoft Access là một thành phần của bộ Tin học văn phòng của Tập đoàn phần mềm lớn nhất thế giới Microsoft. Đây là Hệ quản trị cơ sở dữ liệu quan hệ rất mạnh. Có giao diện thân thiện, dễ sử dụng.

4. Mục đích của Quản ly đoàn viên?
Quản lý đoàn viên là ta quản lý các thông tin về đoàn viên là sinh viên trong các trường Đại học
 Các thao tác trong quá trình Quản lý là ta phải lưu hồ sơ về Đoàn viên, tạo mới, sửa đổi, bổ sung, loại bỏ các thông tin hoặc hồ sơ Đoàn viên. Các tiện ích hỗ trợ người quản lý mà chương trình Quản lý Đoàn viên phải có là : Tìm kiếm, Báo cáo,Tính toán, In ấn , …


CHƯƠNG II. PHÂN TÍCH CƠ SỞ DỮ LIỆU
Đặt vấn đề :
Khi Quản lý đoàn viên-sinh viên ta phải Quản lý thông tin gì về Đoàn viên?
Giải quyết
-	Ta phải quản lý thông tin về đơn vị công tác : Trường , Liên Chi đoàn , Chi đoàn
-	Thông tin cá nhân : Má số sinh viên , Họ tên , ngày sinh , quê quán , nơi ở hiện nay , giới tính , dân tộc , tôn giáo , 
-	Quản lý đoàn phí 
-	Theo dõi đánh giá hàng năm
-	Thông tin về gia đình
-	Trình độ , năng khiếu , khả năng , chính trị ,…
Để quản lý được những thông tin đó ta dùng các đối tượng sau:
*Phân tích cơ sở dữ liệu 
 Khi quản lý Đoàn viên thì ta phải quản lý theo cái gì ? Thông  thường ta dùng MSSV để quản lý .
Vì sao trong quản lý đoàn viên ta lại quản lý  bằng MSSV ? Vì trong các trường Đại học hoặc Cao đẳng, đối tượng Đoàn viên cần quản lý chính là Sinh viên. Nếu dùng MSSV thì ta sẽ dễ dàng quản lý hơn và có mối liên hệ với các phòng ban khác cũng quản lý cùng một đối tượng là Đoàn viên - Sinh viên thì có thể kết hợp cùng với các phòng ban khác cùng quản lý đối tượng đó. Nó tạo nên tính thống nhất trong quá trình quản lý của cùng một cơ quan, đó là các trường Đại học hay Cao đẳng.
Vấn đề đặt ra là ta phải tổ chức các bảng  và xây dựng chương trình như thế nào cho hợp lý, đảm bảo các tính an toàn và toàn vẹn dữ liệu
+ An toàn dữ liệu : là sự bảo vệ dữ liệu trong Cơ sở dữ liệu chống lại sự truy nhập sửa đổi hoặc phá huỷ bất hợp pháp. Có các cách sau để bảo đảm an toan dữ liệu là: (1) Nhận diện : xác minh người sử dụng hợp pháp (xác minh người được phép hoặc uỷ quyền truy nhập vào CSDL) -> Mật khẩu  PASSWORD, (2) Xác lập khung nhìn (giới hạn phạm vi dữ liệu và quyền truy nhập), (3) Tuyên bố quyền và kiểm tra truy nhập. Trong bài này em chỉ dùng cách thứ nhất : xác minh người truy cập, tức là chỉ một số người được phép truy nhập vào chương trình.
+ Toàn vẹn dữ liệu : là sự bảo vệ chống lại sự truy nhập sửa đổi không có căn cứ bảo vệ 
Các thuộc tính của Đoàn viên là ( MSSV, MSDV, Họ tên, Ngày sinh, giới tính, dân tộc, tôn giáo, quê quán, nơi ở hiện nay, chi đoàn, liên chi đoàn, trường, khoá học, ngày vào đoàn, nơi vào đoàn, ngày cấp thẻ, nơi cấp thẻ, các thông tin về cha, các thông tin về mẹ, trình độ đoàn viên, năng khiếu, lý luận chính trị, ngoại ngữ, ưu khuyết điểm hàng năm, ..)
Ta thấy : rõ ràng MSSV chính là khoá tối thiểu vì chỉ cần biết MSSV ta sẽ xác định được tất cả các thuộc tính còn lại vì MSSV là duy nhất, ngược lại một trong số các thông tin còn lại không thể xác định duy nhất một Đoàn viên vì có thể nhiều đoàn viên có cùng thuộc tính đó.
Ta có :
MSSV -> (MSDV, Chi đoàn, Liên Chi đoàn, Khoá học, Các thuộc tính khác).
MSSV->Gia đình đoàn viên (Họ tên Cha, Nghề nghiệp cha, Họ tên mẹ, nghề nghiệp của mẹ, Tình trạng gia đình, Số con (nếu có gia đình))
MSSV->Trình độ của đoàn viên(Trình độ, Ngoại ngữ, Lý luận chính trị, Năng khiếu, khả năng ).
SSV->Theo dõi hàng năm(Năm, ưu khuyết điểm, Xếp loại)
            Chi đoàn -> Liên chi đoàn
            Chi đoàn -> Khoá học

Lược đồ  quan hệ (trang sau):








![Ảnh chụp màn hình 2024-06-19 013929](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/c8892554-e455-4a40-9be9-ef51e8a8e518)






































          

                                                                                              
Lược đồ trên chưa chuẩn hoá vì các thuộc tính không khoá của quan hệ còn phụ thuộc bắc cầu vào khoá chính .
Ta phân tích thành như sau để đưa về dạng chuẩn hoá


		



![Ảnh chụp màn hình 2024-06-19 014238](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/14f8a7ac-4f76-43b7-af8a-3faf677aba19)









 	
 
Như vậy ta có thể chia thành một số bảng chính sau:
Lý lịch đoàn viên ( MSSV,Chi đoàn, Liên chi đoàn, Khoá, Các thông tin khác (Mã ĐV, Họ tên, ngày sinh, giới tính, dân tộc, tôn giáo, quê quán, chỗ ở hiện tại, )) : Chứa lý lịch của Đoàn viên.
Gia đình đoàn viên (MSSV, Họ tên cha, nghề nghiệp cha, họ tên mẹ, nghề nghiệp mẹ , tình trạng gia đình, số con (nếu có gia đình))
Trình độ đoàn viên (MSSV, Trình độ, ngoại ngữ, lý luận chính trị, năng khiếu, sở trường ).
Theo dõi hàng năm (MSSV, năm, Ưu khuyết điểm, Xếp loại)
Chi đoàn –Liên Chi đoàn (Liên chi đoàn, Chi đoàn) : Chứa các nhóm chi đoàn của các khoa.
Chi đoàn-Khoá học (Chi đoàn, Khoá học) : Chứa các nhóm chi đoàn của các Khoá học khác nhau.

 Và còn một số bảng phụ phục vụ cho quá trình tính toán, biểu diễn  số liệu.
Từ phân tích trên ta có thể thiết kế các đối tượng sau


































CHƯƠNG III. CÁC ĐỐI TƯỢNG

I. Đối tượng bảng
1.Các bảng 
    Có tổng cộng 15 bảng  như trong Hình1 dưới đây :
     


![Ảnh chụp màn hình 2024-06-17 092857](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/16a0bdcf-cb48-48e5-a9db-cb2547199975)









    
Tên trường	Kiểu dữ liệu	Giải thích
MA SV	Number	Mã số sinh viên của Đoàn viên
MSDV	Number	Mã số Đoàn viên của đoàn viên
HO DEM	Text	Họ và tên đệm của đoàn viên
TEN	Text	Tên của đoàn viên
GIOI TINH	Yes/No	Giới tính của đoàn viên
DAN TOC	Number	Mã dân tộc của đoàn viên
TON GIAO	Number	Mã tôn giáo của đoàn viên
QUE QUAN	Text	Quê quán cảu đoàn viên
NOI O HIEN NAY	Text	Nơi ở hiện nay của đoàn viên
NGAY VAO DOAN	Date/Time	Ngày đoàn viên vào đoàn
NOI VAO DOAN	Text	Nơi đoàn viên vào đoàn
NGAY CAP	Date/Time	Ngày cấp thẻ đoàn viên
NOI CAP	Text	Nơi cấp thẻ đoàn viên
CHI DOAN	Number	Mã chi đoàn nơi đoàn viên học tập
LIEN CHI DOAN	Number	Mã Liên chi đoàn nơi đoàn viên học tập
DONG DOAN PHI	Yes/No	đoàn viên đóng đoàn phí hay chưa
KHOA	Number	Mã khoá học của đoàn viên
CHUC DANH CAP TRUONG	Number	Mã chức danh cấp trường của đoàn viên
CHUC DANH CAP LCD	Number	Mã chức danh cấp liên chi đoàn của đoàn viên
CHUC DANH CAP CD	Number	Mã chức danh cấp Chi đoàn của đoàn viên
TRUONG	Number	Mã trường của đoàn viên

Bảng 2 : 			T_Dan toc
  Bảng này gồm 2 trường, trường khoá là : MA DT. Trường này để lưu mã của dân tộc, còn trường DAN TOC để lưu tên của dân tộc. 
![Ảnh chụp màn hình 2024-06-17 085426](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/acd9f5ec-8d3f-447b-a113-f36ca2aa5dd3)
Hình 2 – Bảng dân tộc

Bảng 3:		T_Ton giao
Bảng này gồm 2 trường , trường khoá là MA TON GIAO để lưu mã của tôn giáo còn trường TON GIAO để lưu tên của tôn giáo tương ứng
 ![Ảnh chụp màn hình 2024-06-17 085808](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/0a781752-bc95-4c36-ba2a-0e37ccfac47c)
Hình 3 – Bảng tôn giáo

Bảng 4:		T_Gioi tinh
 Trường MA GT có kiểu Yes/No để lưu mã của giới tính là trường khoá, còn trường GIOI TINH để lưu tên của giới tính tương ứng
 ![Ảnh chụp màn hình 2024-06-17 090017](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/c43c6e44-55d2-434b-874e-41f310bda908)
Hình 4 – Bảng giới tính

Bảng 5:		T_Chuc danh
Bảng này dùng để lưu các 
 ![Ảnh chụp màn hình 2024-06-17 090122](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/44f1ed05-7c83-4342-a7a7-ecaa21fec4a1)
Hình 5 – Bảng chức danh

Bảng 6:		T_Trinh do DV
Bảng này lưu trữ các thông tin về trình độ văn hoá và lý luận chính trị của đoàn viên   có các trường và kiểu dữ liệu như trong hình
 ![Ảnh chụp màn hình 2024-06-17 090459](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/8798bb08-7641-4f82-a5eb-a4fe7b610e04)
Hình6 – Bảng trình độ đoàn viên

Bảng 7:		T_Chi doan_Khoa
Bảng này lưu trữ các chi đoàn trong cùng một khoá của tất cả các khoa
 ![Ảnh chụp màn hình 2024-06-17 091003](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/6d418bfb-702e-4cd5-a246-6ed7863a8b7e)
Hình7 – Bảng trình chi đoàn theo khoá

Bảng 8:		T_Khoa hoc
Bảng này lưu trữ các khoá học của trường , kết quả này ta không nhập trực tiếp mà chương trình tự động tính toán khi biết khoá hiện tại và số năm đào tạo của trường 
 ![Ảnh chụp màn hình 2024-06-17 091109](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/d7bc235d-fdec-4d8e-b3b4-58cdb3b2412e)
Hình 8– Bảng khoá học

Bảng 9:		T_Truong
Bảng này lưu các thông tin về trường mà ta cần quản lý. Các thông tin cơ bản là : tên trường, khoá hiện tại, năm hiện tại, số năm đào tạo.
 ![Ảnh chụp màn hình 2024-06-17 091305](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/caf79850-026b-4545-8c06-b25e7391d2c0)
Hình 9 – Bảng trường

Bảng 10:		T_Lien chi doan
Bảng này lưu thông tin về các Liên chi đoàn trực thuộc trường
 ![Ảnh chụp màn hình 2024-06-17 091405](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/f4398521-8807-47bf-a41a-a07367161f66)
Hình 10 – Bảng Liên chi đoàn

Bảng 11			T_Lien chi doan_Chi doan
Bảng này lưu các chi đoàn tất cả các khoá học của cùng một Liên chi đoàn 
 ![Ảnh chụp màn hình 2024-06-17 091534](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/a449946d-d25f-4575-a6ab-cb9475148edb)
Hình 11 – Bảng chi đoàn của một Liên chi đoàn

Bảng 12:		T_Gia dinh doan vien
Bảng này lưu thông tin về gia đình của đoàn viên, các thông tin đó là về cha mẹ, tình trạng gia đình (có vợ hoặc chồng, số con nếu có).
 ![Ảnh chụp màn hình 2024-06-17 091858](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/9912937e-7070-4b98-b0e2-828f560998a9)
Hình 12 – Bảng gia đình đoàn viên

Bảng 13:			T_Danh gia hang nam
Bảng này lưu các thông tin về theo dõi ưu khuyết điểm hàng năm của đoàn viên 
 ![Ảnh chụp màn hình 2024-06-17 092133](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/603f968f-0cfa-45dd-bfdc-084895358c10)
Hình 13 – Bảng theo dõi hàng năm của đoàn viên

Bảng 14:		T_Mat khau
Bảng này lưu thông tin về tên ngướiử dụng chương trình và mật khẩu 
 ![Ảnh chụp màn hình 2024-06-17 092436](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/9565524f-5d91-40c4-a4e9-665b54297514)
Hình 14 – Bảng mật khẩu

Bảng 15:		T_Xep loai
Bảng này lưu các loại của xếp loại : có 4 loại ( Xuất sắc, khá, trung bình, yếu kém).
 ![Ảnh chụp màn hình 2024-06-17 092527](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/97067ba4-f414-43c1-a866-7d5231911d87)
Hình 15– Bảng xếp loại

2. Các quan hệ
    Quan hệ giữa các bảng được thể hiện như trong Hình16  sau:
      ![Ảnh chụp màn hình 2024-06-19 015800](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/100f7f12-db2a-446a-947d-6ca5e787f244)
Hình16 – Quan hệ giữa các bảng

II . Đối tượng Form
 Form là thành phần chủ yếu để thiết kế phục vụ cho nhập dữ liệu, xuất dữ liệu và thực hiện các thao tác trên Cơ sở dữ liệu
Danh sách các Form trong bài này được thể hiện như trong hình sau:
 ![Ảnh chụp màn hình 2024-06-19 020014](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/ad1d8413-0298-48f3-8f69-4e74250fe032)
Hình 17 – Danh sách các bảng

*Diễn giải từng Form cụ thể :

Form1	 F_Thong tin ban dau
 Đây là form đầu tiên xuất hiện khi ta bắt đầu sử dụng chương trình. Form này chỉ xuất hiện một lần trong suốt quá trinh sử dụng chương trình. Ta phải nhạp vào các thông tin ban đầu về : Tên trường, Năm học, Khoá học tương ứng của năm đó, Số năm đào tạo của trường, Nhập tên người quản lý, mật khẩu của người đó. Khi đã nhập một lần thì đến lần thứ hai, chương trình không mở form nay nữa mà sẽ mở form Đăng ký truy nhập.
 ![Ảnh chụp màn hình 2024-06-19 020144](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/a3e1babf-6eff-405b-881a-5d140f2e4724)
Hình 18 – form Thông tin ban đầu

   Form2	 F_Dang ky truy nhap
 Form này sẽ xác nhận người truy nhập có nằm trong danh sách những người được phép truy nhập vào chương trình hay không.
Khi xuất hiện Form này ta phải nhập tên đăng ký và mật khẩu tương ứng. Nếu nhập sai sẽ không thể nào truy nhập được vào chương trình mà phải thoát về Windows. Đây là một cách đảm bảo tính an toàn dữ liệu.
 ![Ảnh chụp màn hình 2024-06-19 020245](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/80523c9e-7572-47cb-9b82-6330fc00c3dc)
Hình19 – form F_Mat khau

 Form3	 F_Chuong trinh chinh
Form chính để gắn hệ thống menu và có các công cụ để lựa chọn nhanh thay cho việc chọn từ các menu. Nó còn thể hiện được tên của trường, năm học tương ứng, bản quyền. Giao diện chính của Form chính như hình sau:










![image](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/192585b9-1bea-441f-8c6b-e8d1f4dc908a)













 
Hình 20 – Giao diện của Form chính
       (1)Là hệ thống menu
      (2)là các nút lệnh để lựa chọn nhanh nó sẽ gọi đến các form khác
        (3)là tên trường và năm học tương ứng
	
Form4    F_Nhap du lieu
Là form xuất hiện khi ta click vào nút Nhập DL. Khi xuất hiện nó có giao diện như sau.
 ![image](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/f8d1df2d-b7e7-4958-b18e-a0dcb5993490)
Hình 21– form F_nhap du lieu

 Tại đây ta sẽ có các lựa chọn giống như khi ta click vào menu Nhap du liệu. Click vào nút bất kỳ ta có các form để thực hiện các công việc tương ứng
 
Form5	 F_Tien ich
Khi click vào nút Tiện ích trên Form chính ta thấy sẽ xuất hiện Form này có giao diện như sau. Trên Form này cũng có các nút lệnh tương ứng như khi ta click vào menu Tien ich trên Form chinh
Mỗi lựa chọn sẽ cho xuất hiện form tương ứng với công việc tương ứng .
 ![image](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/67441ba8-ed01-4fc0-8a09-c53958cd753f)
Hình 2 2 – Form F_tien ich

 Form 6 	F_Danh muc
 form này cho các lựa chọn để xem danh sách 
 ![Ảnh chụp màn hình 2024-06-19 020638](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/b69cf727-d199-4c8d-a029-4e23b87fea87)
Hình 2 3 – Form F_Danh muc

 Form 7 	frmTimkiem
 Form này có giao diện như trong hình sau . Nhiệm vụ của nó là phải tìm kiếm Đoàn viên theo điều kiện nào đó . ở đây có  7 trường tìm kiếm cùng với điều kiện And hay OR  . Ta có thể kết hợp nhiều trường cùng lúc để tìm kiếm . Điều này sẽ tạo điều kiện thuận lợi khi ta muốn tìm kiếm ai đó mà không nhớ MSSV mà chỉ nhớ một trong các thông tin khác của Đoàn viên cần tìm . ở đay khi tìm kiếm có dùng ký tự thay thế * cho nhiều ký tự . Chẳng hạn muốn tìm Đoàn viên có mã bắt đầu bằng số 1 ta gõ vào trong trường tìm kiếm MSSV là 1* (như hình ) sẽ có kết quả các Đoàn viên có mã sinh viên bắt đầu bằng số1 (như hình )
 ![Ảnh chụp màn hình 2024-06-19 020731](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/cd5559ae-2308-4d5b-81c7-ff659dc662ad)
Hình 24 – form frmTimkiem

 Form 8	F_Bao cao
form này cho 2 lựa chọn để xem báo cáo và xem thống kê
 ![Ảnh chụp màn hình 2024-06-19 020810](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/4b5b67ac-8851-41a0-b760-e8ae1d447790)
Hình 25 – Form F_Bao cao

 Form 9	 F_Thong ke
Form này cho các lựa chọn thống kê về dân tộc , tôn giáo , giới tính , đặc biệt là tình hình đoàn phí và Danh sách Đoàn viên chưa đóng Đoàn phí
Giao diện chính là:
 ![Ảnh chụp màn hình 2024-06-19 020858](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/983c02a6-bfd9-4908-8ecb-cf1979dc5ca9)
Hình 26 – Form F_Thong ke

 Form 10 	F_In an
Form này cho các lựa chọn để in ấn  
 ![Ảnh chụp màn hình 2024-06-19 020932](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/78952e95-2182-4b3f-a50b-cc5d11f090e6)
Hình 27 -  Form 	F_In an

Form 11	 F_Ly lich doan vien
form này có tác dụng để nhập , xem ,sửa ,xoá hồ sơ đoàn viên .
 ![Ảnh chụp màn hình 2024-06-19 021018](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/166272d3-8202-4860-83ab-89378d2d70ab)
Hình 28 – form  F_Ho so doan vien

Form 12       F_Danh sach Lien chi doan
Form này dùng để xem và nhập thêm Liên chi đoàn   
![image](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/2d0d98ac-63ee-4682-b83a-ae38620e99af)
Hình 29 – form F_Danh muc Lien chi doan
 
Form 13      		F_Nhap chi doan
Form này dùng để nhập thêm chi đoàn của các Khoa và các khoá 
 ![Ảnh chụp màn hình 2024-06-19 021219](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/f7faf28b-2a1c-4252-b555-d49e7441249b)
Hình 30- form F_Nhap chi doan

Form 14		F_Them nguoi su dung
 Form này dùng để nhập thêm người sử dụng chương trình
 ![Ảnh chụp màn hình 2024-06-19 021255](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/dcebfe86-fe05-40f3-b1dc-760a16adf48a)
Hình 31 - F_Them nguoi su dung

Form 15		F_Thay mat khau
Form này dùng để thay đổi mật khẩu cho người quản lý
![Ảnh chụp màn hình 2024-06-19 021330](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/23d0c708-23f2-4351-8add-2e4e947ebf77)
Hình 32 – form F_Thay mat khau

Form	16	F_Cap nhat thong tin ve truong
		form này dùng để thay đổi các thông tin về trường 
 ![image](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/98d93d1b-c039-4362-b379-11b0aaecc8e7)
Hình 33 - F_Cap nhat thong tin ve truong

Ngoài các form trên còn một số form sau:


III . Macro và hệ thống menu
          1. Macros
Các macros tronh bài này được sử dụng chỉ với mục để tạo ra hệ thống menu 
![image](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/bde8903f-f903-4a51-873d-6eba9ab6fcde)

     2 . Hệ thống menu 
               Menu cấp 0
 
![Ảnh chụp màn hình 2024-06-19 021547](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/f42e2414-75de-4cd1-be88-1617f900c565)

a. Menu Nhap du lieu
 ![Ảnh chụp màn hình 2024-06-19 021634](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/c865c59e-95c7-4c09-9713-6fbd03d725a6)
 
b.Menu Tien ich
![Ảnh chụp màn hình 2024-06-19 021709](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/085d42fc-79af-4e18-9f47-6064ae0dce39)                              
 
c.Menu Danh sach
 ![Ảnh chụp màn hình 2024-06-19 021754](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/1d2b4a8c-70ac-4a8f-ab3d-35ae35406acd)

d.Menu Timkiem
![Ảnh chụp màn hình 2024-06-19 021826](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/9b859838-fdbe-4f1e-b4c7-b610b01efba4)

e.Menu Bao cao
![Ảnh chụp màn hình 2024-06-19 021902](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/b499d734-9168-40f4-800a-5b9a1e8f7baf)

f.Menu In an
![Ảnh chụp màn hình 2024-06-19 021932](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/7b938c16-657e-4784-a176-7d9715a75116)

g. Menu Thoat        

![Ảnh chụp màn hình 2024-06-19 022000](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/48e03591-9e7a-48dd-ad11-688a0cb474f5)

                                             
IV .Đối tượng  Querrry
Hệ thống Querry là 
 
![Ảnh chụp màn hình 2024-06-19 022047](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/5684a8ec-380b-4bb4-bac3-ea4adda8ea31)

Q_Bao cao chi doan : Truy vấn chọn xem danh sách đoàn viên đã xếp loại . Phục vụ cho Report R_Bao cao chi doan

 ![Ảnh chụp màn hình 2024-06-19 022133](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/eb2d1e9e-969d-4b57-8afd-9e52d12512ab)


Q_Chi doan theo LCD : Truy vấn các chi đoàn của một Liên chi đoàn 
 
![Ảnh chụp màn hình 2024-06-19 022208](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/3cfe72ac-8dc9-4aa6-9032-839f4f85d3c1)

Q_Danh sach BCH cap truong : truy vấn danh sách đoàn viên nằm trong ban chấp hành cấp trường

Q_Danh sach BCH LCD : truy vấn danh sách đoàn viên nằm trong ban chấp hành các Liên chi đoàn
 
![Ảnh chụp màn hình 2024-06-19 022249](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/280a186b-5863-4fb3-87a7-ecbdb0c3f9fe)

Q_Danh sach BCH CD : truy vấn danh sách đoàn viên nằm trong ban chấp hành của các chi đoàn
Q_Danh sach doan vien mot LCD : truy vấn chọn danh sách đoàn viên của một Liên chi đoàn
Q_Danh sach Dv dan toc thieu so : truy vấn chọn danh sách đoàn viên là dân tộc thiểu số
Q_DS DV chua dong doan phi : truy vấn chọn ra danh sách đoàn viên chưa đóng đoàn phí

 ![Ảnh chụp màn hình 2024-06-19 022319](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/d9b27233-4081-4924-b7ea-deb3f8665a64)

Q_Ho so doan vien : truy vấn các thông tin đầy đủ về một đoàn viên
Q_Ly lich doan vien : truy vấn các thông tin tổng quát về một đoàn viên
Q_Ly lich doan vien mot CD : truy vấn chọn ra danh sách đoàn viên cùng một số thông tin tổng quát của một chi đoàn 
Q_thong ke theo dan toc : truy vấn tính tổng tìm ra tổng số đoàn viên của các dân tộc khác nhau

 ![Ảnh chụp màn hình 2024-06-19 022359](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/16ed1bb1-48af-4f54-bc19-0e07469a4f55)

Q_thong ke theo ton giao : truy vấn tính tổng tìm ra tổng số đoàn viên của các tôn giáo khác nhau
 ![Ảnh chụp màn hình 2024-06-19 022432](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/aaf7b840-f9a1-4c65-a17e-0fa81254f304)

Q_Thong ke theo gioi tinh : truy vấn tính tổng tìm ra tổng số đoàn viên của nam và nữ
 ![Ảnh chụp màn hình 2024-06-19 022506](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/0cfe371e-94de-4888-9cb9-66e4de0d6f44)

Q_Thong ke tinh hinh doan phi  : truy vấn tính tổng tìm ra tổng số đoàn viên chưa nộp đoàn phí
 ![Ảnh chụp màn hình 2024-06-19 022533](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/36bdae01-68db-427b-aa9f-d2e64e0f115d)

Q_Thong ke xep loai cua mot CD : truy vấn để báo cáo tình hình xếp loại của một chi đoàn trong một năm
 ![Ảnh chụp màn hình 2024-06-19 022604](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/df73e8b3-0d71-4a9e-8546-405039e48703)





V .Đối tượng  report
Các report là công cụ chủ yếu để thống kê , báo cáo , in ấn , trong bài này em dùng 16 report sau:
 ![Ảnh chụp màn hình 2024-06-19 022639](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/eca8019f-62b5-49d2-9ebd-b986986a5b6e)

R_Danh sach BCD LCD : để báo cáo và in ấn danh sách ban chấp hành cấp liên chi đoàn
R_Thong ke xep loai cua mot CD : dùng báo cáo hàng năm cho một chi đoàn về tình hình đoàn viên và tỷ lệ các loại mà chi đoàn dạt được
R_Bao cao chi doan : 
R_Danh sach doan vien
R_Danh sach doan vien mot LCD
R_Danh sach DV dan toc thieu so
R_Danh sach doan vien mot chi doan : dùng để xem và in ra danh sách đoàn viên  một chi đoàn 
R_DS doan vien chua nop doan phi : dùng báo cáo hàng năm  về tình đóng đoang phí của đoàn viên
R_Ho so doan vien : dùng để in hồ sơ cho một đoàn viên
R_Thong ke theo dan toc : dùng báo cáo hàng năm  về tỷ lệ các dân tộc
R_Thong ke theo ton  giao : dùng báo cáo hàng năm  về tỷ lệ các tôn giáo
R_Thong ke tình hình doan phi : dùng báo cáo hàng năm  về tỷ lệ đoàn viên chưa nộp đoàn phí





Chương 4
Sử dụng chương trình - Ví dụ



Cách dễ hiểu nhất và rõ ràng nhất để biết một phần mềm làm việc như thế nào là thông qua Ví dụ sử dụng chương trình . Vì lý do đó em kết hợp “Ví dụ” và “Cách sử dụng chương trình” . 

Muốn quản lý Đoàn viên của trường Đại học Xây dựng , ta làm các bước sau :
1. Nhập các thông số ban đầu 
  Khi bắt đầu khởi động chương trình sẽ yêu cầu nhập các thông số ban đầu , giả sử ta nhập  như sau (Mật khẩu là :123456)
 ![Ảnh chụp màn hình 2024-06-19 022715](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/4b17449e-3d5e-4802-9dd1-9761974b3907)

Sau khi nhập xong , nhấn nút Chấp nhận chương trình sẽ hiển thị Form chính có giao diện như  hình 21 . Sau khi nhập xong , chương trình sẽ tự động tính cho ta các khoá học của trường trong năm đó . Chẳng hạn ta sẽ có được các khoá học là : 48,47,46,45,44

2. Nhập số liệu
2.1.	Nhập danh sách các khoa
 Từ Form chính để nhập danh sách các Liên chi đoàn ta có 2 cách sau :
Cách 1: Click nút Nhập dữ liệu trên Form , khi đó sẽ xuất hiện form Nhập dữ liệu như hình sau . 
 ![Ảnh chụp màn hình 2024-06-19 022826](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/a36a6a76-dade-4188-a8ec-ba04cff6626b)

Từ đây ta sẽ chọn nút  “Nhập DS Liên Chi Đoàn” để gọi Form Nhập dữ liệu cho Liên chi đoàn như hình 28 . Tại đây ta sẽ Nhập các Liên chi đoàn và mã tương ứng .Sau khi nhập xong ta click nút “Thoát”
Cách 2 : Chọn từ menu  Nhap du lien->Nhap du lieu->Nhap DS Lien chi doan . Ta sẽ nhập như trên 

         2.2  Nhập danh sách các chi đoàn
 Ta cũng có 2 cách
Cách 1:
Ta sẽ chọn nút  “Nhập DS  Chi Đoàn” để gọi Form Nhập dữ liệu cho Chi đoàn như hình 29 . Để nhập danh sách chi đoàn cho một khoa của một khoá ta phải chọn khoa và khoá tương ứng rồi đánh tên của chi đoàn cần nhập vào textbox Chi đoàn . Sau đó click nút nhập . Để nhập chi đoàn tiếp theo ta đánh lại tên chi đoàn cần nhập vào textbox chi đoàn  .Sau khi nhập xong ta click nút “Thoát”
Cách2 : 
Chọn từ menu  Nhap du lieu->Nhap du lieu->Nhap DS Chi doan . Ta sẽ nhập như trên 
     2.3 .Nhập hồ sơ đoàn viên cho từng chi đoàn của từng khoa
 Sau khi đã nhập xong danh sách Liên chi đoàn và chi đoàn thì ta mới bắt đàu nhập hồ sơ đoàn viên cho từng chi đoàn .
Để mở được Form nhập ta có 2 cách sau
Cách 1 : Chọn từ menu “Nhap du lieu ->Nhap du lieu->Nhap ho so Doan vien
Cách 2 : Chọn nút  “Thêm hồ sơ Đoàn viên” như hình trên 
Khi Form Nhập hồ sơ này xuất hiện , ta nhập các thông tin vào , không nhất thiết phải nhập đầy đủ thông tin nhưng một số trường bắt buộc phải nhập (MSSV , Khoá , Chi đoàn , Liên chi đoàn , Dân tộc , Tôn giáo , Giới tính)
Sau khi nhập xong một đoàn viên ,  muốn nhập tiếp ta chọn nút Thêm trên form này 
Đến đây ta đã nhập xong dữ liệu cho chương trình . 
3. Xem dữ liệu
 Khi muốn xem hồ sơ Đoàn viên
      Có 2 cách 
Cách 1: Chọn từ menu Nhap du lieu->Xem ho so Doan vien
Cách 2: Chọn từ nút Xem ho so doan vien từ form Hồ sơ (hình?)
Trong quá trình xem ta dùng các nút “Tiếp theo” và ‘Quay lại” để xem hồ sơ tiếp theo hoặc trước đó
 Khi muốn sửa hồ sơ ta chọn nút “Sửa” và sửa , sau đó nhấn nút “ Lưu ” . 
4.Tìm kiếm 
              Khi muốn tìm kiếm một đoàn viên : 
-Nếu biết MSSV của đoàn viên cần tìm thì chỉ cần gõ MSSV vào textbox “MSSV” rồi nhấn phím “Tìm kiếm” ta sẽ có kết quả . Ví dụ : Ta gõ MSSV là 14069 ta sẽ có kết quả
 ![Ảnh chụp màn hình 2024-06-19 022917](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/7b542d4d-33c0-44d9-a87e-eafc3189a538)

Trong trường hợp không biết rõ MSSV ta có thể kết hợp nhiều trường để cùng tìm kiếm . Ví dụ : ta gõ MSSV=1* ,chọn giá trị And trong combo box trước trường tên và nhập tên là “Sứng” thì ta có kết quả như trên . nếu muốn tìm tất cả các đoàn viên có mã  bắt đầu bằng số 1 ta gõ 1* trong trường MSSV ta có kết quả sau:
 ![Ảnh chụp màn hình 2024-06-19 022957](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/396500c2-1fa5-4b57-97e4-5498210614f7)


5.  Báo cáo và thống kê
+Xem Báo của một chi đoàn theo năm : ta có 2 cách 
Cách 1 : chọn trực tiếp từ menu Bao cao->Bao cao theo chi doan
Cách 2 : click nút “Báo cáo” từ form chính -> chọn nút “Báo cáo theo chi đoàn”
 ![Ảnh chụp màn hình 2024-06-19 023032](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/dcf3c530-d896-417f-a269-ad39ae6ef45d)

Sau đó ta sẽ có kết quả 
+Xem thống kê , khi click vao nut “Thống kê” ở hình trên , ta co các lựa chọn sau
 ![Ảnh chụp màn hình 2024-06-19 023117](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/e900737f-cfa6-4e53-8222-6d8788ed2ae0)

chọn một kiểu sau đó nhấn nút chấp nhận  ta có kết quả . Giả sử ta chọn “Tình hình đoàn phí toàn trường” ta có kết quả sau:
6. In ấn 
Để có thể in ấn có thể làm theo một trong 2 cách sau 
Cách 1 : chọn trực tiếp từ menu  In an . Nhưng khi chọn chương trình chỉ cho xem trước , muốn in thì phải chọn
 print  trên thanh công cụ chuẩn
 ![Ảnh chụp màn hình 2024-06-19 023229](https://github.com/0337211793/B-i-t-p-l-n-HQTCSDL/assets/169513994/c56538d8-fc2a-4ccb-b6b5-804b979c1752)


Cách 2 : chọn nút “In ấn” từ chương trình chính . Khi form In ấn xuất hiện (như hình 27) . Ta chọn lấy 1 kiểu . nếu muốn xem trước khi in thì ta chọn nút”Xem trước” nếu muốn in thì chọn nút “In”
Vì report cho ta khả năng In ấn nên ta có thể in ấn bất cứ một báo cáo hay thống kê nào , ta chỉ việc chọn công cụ Print của report 

5.Báo cáo và Thống kê





