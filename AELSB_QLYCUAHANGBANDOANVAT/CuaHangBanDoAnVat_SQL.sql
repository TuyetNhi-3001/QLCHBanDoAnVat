CREATE DATABASE CuaHangBanDoAnVat_SQL
GO
USE CuaHangBanDoAnVat_SQL
GO
--
--
CREATE TABLE CUAHANG
(
	MaCH varchar(25) constraint CH_MaCH_PK primary key,
	SDT varchar(15) not null,
	DiaChi nvarchar(200) not null
)
GO
CREATE TABLE CHUCVU
(
	MaCV varchar(25) constraint CV_MaCV_PK primary key,
	TenCV nvarchar(100) not null,
	MucLuong decimal not null check (MucLuong > 0)
)
GO
CREATE TABLE NHANVIEN
(
	MaNV varchar(25) constraint NV_MaNV_PK primary key,
	TenNV nvarchar(100) not null,
	NamSinh date not null,
	SDT varchar(15) not null,
	DiaChi nvarchar(200),
	NgayNhanViec date,
	MaCV_NV varchar(25),
	MaCH_NV varchar(25) not null,
	constraint NV_MaCVNV_FK foreign key (MaCV_NV) references CHUCVU(MaCV),
	constraint NV_MaCHNV_FK foreign key (MaCH_NV) references CUAHANG(MaCH)
)
GO
CREATE TABLE NHACUNGCAP
(	
	MaNCC varchar(25) constraint NCC_MaNCC_PK primary key,
	TenNCC nvarchar(100) not null,
	SDT varchar(15) not null,
	DiaChi nvarchar(200)
)
GO
CREATE TABLE SANPHAM 
(
	MaSP varchar(25) constraint SP_MaSP_PK primary key,
	TenSP nvarchar(200) not null,
)
GO
CREATE TABLE PHIEUNHAPHANG
(
	MaPN varchar(25) constraint PNH_MaPN_PK primary key,
	NgayNhap date,
	MaNCC_PN varchar(25),
	MaCH_PN varchar(25),
	constraint PNH_MaNCCPN_FK foreign key (MaNCC_PN) references NHACUNGCAP(MaNCC),
	constraint PNH_MaCHPN_FK foreign key (MaCH_PN) references CUAHANG(MaCH)
)
GO
CREATE TABLE CHITIET_PHIEUNHAP
(
	MaPN_CTPN varchar(25),
	MaSP_CTPN varchar(25),
	GiaNhap decimal not null check (GiaNhap > 0),
	SoLuong int not null check (SoLuong > 0),
	ThanhTien decimal,
	NSX date,
	HSD date,
	constraint CTPN_MaPNCTPN_FK foreign key (MaPN_CTPN) references PHIEUNHAPHANG(MaPN),
	constraint CTPN_MaSPCTPN_FK foreign key (MaSP_CTPN) references SANPHAM(MaSP),
	constraint CTPN_MAPN_SP_PK primary key (MaPN_CTPN, MaSP_CTPN)
)
GO
CREATE TABLE KHACHHANG
(
	MaKH varchar(25) constraint KH_MaKH_PK primary key,
	TenKH nvarchar(100) not null,
	SDT varchar(15) not null
)
GO
CREATE TABLE DONHANG
(
	MaDH varchar(25) constraint DH_MaDH_PK primary key,
	NgayDatHang date,
	DiaChi nvarchar(200) not null,
	MaKH_DH varchar(25) not null,
	MaNV_DH varchar(25) not null,
	MaCH_DH varchar(25)not null,
	constraint DH_MaKHDH_FK foreign key (MaKH_DH) references KHACHHANG(MaKH),
	constraint DH_MaNVDH_FK foreign key (MaNV_DH) references NHANVIEN(MaNV),
	constraint DH_MaCHDH_FK foreign key (MaCH_DH) references CUAHANG(MaCH)
)
GO
CREATE TABLE CHITIET_DONHANG
(
	MaDH_CTDH varchar(25),
	MaSP_CTDH varchar(25),
	GiaBan decimal not null check (GiaBan > 0),
	SoLuong int not null check (SoLuong > 0),
	ThanhTien decimal,
	ThanhToan bit,
	constraint CTDH_MaDHCTDH_FK foreign key (MaDH_CTDH) references DONHANG(MaDH),
	constraint CTDH_MaSPCTDH_FK foreign key (MaSP_CTDH) references SANPHAM(MaSP),
	constraint CTDH_MADH_SP_PK primary key (MaDH_CTDH, MaSP_CTDH)
)
GO

CREATE TABLE TAIKHOAN
(
	SDT_DN varchar(15) constraint TK_SDT_PK primary key,
	MatKhau nvarchar(50),
	Quyen varchar(30)
)
GO
--
--
CREATE ROLE Chu
GRANT SELECT, INSERT, UPDATE, DELETE on CHITIET_DONHANG to Chu
GRANT SELECT, INSERT, UPDATE, DELETE on CHITIET_PHIEUNHAP to Chu
GRANT SELECT, INSERT, UPDATE, DELETE on CHUCVU to Chu
GRANT SELECT, INSERT, UPDATE, DELETE on CUAHANG to Chu
GRANT SELECT, INSERT, UPDATE, DELETE on DONHANG to Chu
GRANT SELECT, INSERT, UPDATE, DELETE on KHACHHANG to Chu
GRANT SELECT, INSERT, UPDATE, DELETE on NHACUNGCAP to Chu
GRANT SELECT, INSERT, UPDATE, DELETE on NHANVIEN to Chu
GRANT SELECT, INSERT, UPDATE, DELETE on PHIEUNHAPHANG to Chu
GRANT SELECT, INSERT, UPDATE, DELETE on SANPHAM to Chu
GRANT SELECT, INSERT, UPDATE, DELETE on TAIKHOAN to Chu
GRANT EXEC to Chu

CREATE ROLE QuanLy
GRANT SELECT, INSERT, UPDATE, DELETE on CHITIET_DONHANG to QuanLy
GRANT SELECT, INSERT, UPDATE, DELETE on CHITIET_PHIEUNHAP to QuanLy
GRANT SELECT, INSERT, UPDATE, DELETE on CHUCVU to QuanLy
GRANT SELECT, INSERT, UPDATE, DELETE on CUAHANG to QuanLy
GRANT SELECT, INSERT, UPDATE, DELETE on DONHANG to QuanLy
GRANT SELECT, INSERT, UPDATE, DELETE on KHACHHANG to QuanLy
GRANT SELECT, INSERT, UPDATE, DELETE on NHACUNGCAP to QuanLy
GRANT SELECT, INSERT, UPDATE, DELETE on NHANVIEN to QuanLy
GRANT SELECT, INSERT, UPDATE, DELETE on PHIEUNHAPHANG to QuanLy
GRANT SELECT, INSERT, UPDATE, DELETE on SANPHAM to QuanLy
GRANT SELECT, INSERT, UPDATE, DELETE on TAIKHOAN to QuanLy
GRANT EXEC to QuanLy

--
--
CREATE PROCEDURE SP_CreateChu 
(
	@TenLogin varchar(15),
	@SdtChu varchar(15),
	@PW varchar(20),
	@Role varchar(30)
)
AS
BEGIN
	EXEC sys.sp_addlogin
		@TenLogin,
		@PW
	EXEC sys.sp_adduser
		@TenLogin,
		@SdtChu
	EXEC sys.sp_addrolemember
		@Role,
		@SdtChu
END;


GO


--
--
EXEC SP_CreateChu
	@TenLogin = '0999999999',
	@SdtChu = '0999999999',
	@PW = '0999999999',
	@Role = 'Chu'

EXEC SP_CreateChu
	@TenLogin = '0961012560',
	@SdtChu = '0961012560',
	@PW = '0961012560',
	@Role = 'QuanLy'
--
--
INSERT INTO CUAHANG VALUES ('CH01', '0111111111', '509 Trần Xuân Soạn, phường Tân Kiểng, Quận 7, TP HCM')
INSERT INTO CUAHANG VALUES ('CH02', '0222222222', '280 An Dương Vương, Quận 5, TP HCM')
INSERT INTO CUAHANG VALUES ('CH03', '0333333333', '222 Lê Văn Sỹ, Quận 3, TP HCM')	

INSERT INTO CHUCVU VALUES ('CV01', 'Quản lý cửa hàng')
INSERT INTO CHUCVU VALUES ('CV01', 'Nhân viên bán hàng')

INSERT INTO NHANVIEN VALUES ('NV01', 'Huỳnh Thị Tường Vi', '2004/02/24', '0961012560', 'TP HCM', '2023/10/06', 'CV01', 'CH01')