using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace QLyCHBanDoAnVat
{
    public partial class Form_NguoiQuanLy : Form
    {
        public string SDT_DN;
        private string sqlConn;
        SqlConnection conSql;
        public Form_NguoiQuanLy()
        {
            InitializeComponent();
        }
        public Form_NguoiQuanLy(string conn, string sDT_DN)
        {
            InitializeComponent();
            sqlConn = conn;
            SDT_DN = sDT_DN;
        }

        private void Form_NguoiQuanLy_Load(object sender, EventArgs e)
        {
            lbDate.Text = DateTime.Now.ToString("dd/MM/yyyy");
            conSql = new SqlConnection(sqlConn);
            if (conSql.State == ConnectionState.Closed)
            {
                conSql.Open();
            }

            string query = "SELECT MaSP AS N'Mã sản phẩm', TenSP AS N'Tên sản phẩm' FROM SANPHAM";
            dtgList_Load(dtgListSP, query);
            cbxSanPham_Load();

            /*
            guiMa(cbxMaSP, cbxMaSPNhap);
            guiMa(cbxMaCH, cbxMaCHNhap);
            guiMa(cbxMaCH, cbxMaCHnv);
            guiMa(cbxMaSP, cbxMaSPdh);
            guiMa(cbxMaKH, cbxMaKHdh);
            guiMa(cbxMaCH, cbxMaCHban);
            guiMa(cbxMaNV, cbxMaNVban); */

            if (conSql.State == ConnectionState.Open)
            {
                conSql.Close();
            }
        }

        /***********************/
        /******Dùng chung*******/

        public string MaCH()
        {
            string MaCHnv = "";
            if (conSql.State == ConnectionState.Closed)
            {
                conSql.Open();
            }
            string query = "SELECT MaCH_NV FROM NHANVIEN WHERE SDT = '" + SDT_DN + "' ";
            SqlCommand cmd = new SqlCommand(query, conSql);
            SqlDataReader rd = cmd.ExecuteReader();
            if (rd.Read())
            {
                MaCHnv = rd["MaCH_NV"].ToString();
            }

            if (conSql.State == ConnectionState.Open)
            {
                conSql.Close();
            }
            return MaCHnv;
        }
        public void dtgList_Load(DataGridView dtg, string query)
        {
            DataTable tb = new DataTable();
            if (conSql.State == ConnectionState.Closed)
            {
                conSql.Open();
            }
            SqlCommand cmd = new SqlCommand(query, conSql);
            SqlDataAdapter adt = new SqlDataAdapter();
            adt.SelectCommand = cmd;
            tb.Clear();
            adt.Fill(tb);
            dtg.DataSource = tb;
            dtg.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            if (conSql.State == ConnectionState.Open)
            {
                conSql.Close();
            }
        }
        public bool checkContent(string str)
        {
            if (str != "")
            {
                return true;
            }
            return false;
        }
        private void Textbox_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar);
        }
        public void guiMa(ComboBox cbx1, ComboBox cbx2)
        {
            for (int i = 0; i < cbx1.Items.Count; i++)
            {
                cbx2.Items.Add(cbx1.Items[i].ToString());
            }
        }

        public void thanhTien_PN()
        {
            int soluongnhap = int.Parse(txtSoLuongNhap.Text);
            int gianhap = int.Parse(txtGiaNhap.Text);
            txtThanhTienNhap.Text = (gianhap * soluongnhap).ToString();
        }


        /********NHANVIEN*******/

        public void cbxNhanVien_Load()
        {
            cbxMaCV.Items.Clear();
            cbxTenCV.Items.Clear();
            cbxMaNV.Items.Clear();
            cbxMaCHnv.Items.Clear();

            if (conSql.State == ConnectionState.Closed)
            {
                conSql.Open();
            }
            string query = "SELECT MaCV FROM CHUCVU";
            SqlCommand cmd = new SqlCommand(query, conSql);
            SqlDataReader rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                cbxMaCV.Items.Add(rd["MaCV"].ToString());
            }
            rd.Close();
            query = "SELECT TenCV FROM CHUCVU";
            cmd = new SqlCommand(query, conSql);
            rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                cbxTenCV.Items.Add(rd["TenCV"].ToString());
            }
            rd.Close();
            query = "SELECT MaNV FROM NHANVIEN";
            cmd = new SqlCommand(query, conSql);
            rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                cbxMaNV.Items.Add(rd["MaNV"].ToString());
            }

            if (conSql.State == ConnectionState.Open)
            {
                conSql.Close();
            }
            cbxMaCHnv.Items.Add(MaCH());
        }

        /********SANPHAM*********/

        public void cbxSanPham_Load()
        {
            cbxMaSP.Items.Clear();
            cbxTenSP.Items.Clear();

            if (conSql.State == ConnectionState.Closed)
            {
                conSql.Open();
            }
            string query = "SELECT MaSP FROM SANPHAM";
            SqlCommand cmd = new SqlCommand(query, conSql);
            SqlDataReader rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                cbxMaSP.Items.Add(rd["MaSP"].ToString());
            }
            rd.Close();

            query = "SELECT TenSP FROM SANPHAM";
            cmd = new SqlCommand(query, conSql);
            rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                cbxTenSP.Items.Add(rd["TenSP"].ToString());
            }
            if (conSql.State == ConnectionState.Open)
            {
                conSql.Close();
            }
        }

        /********KHACHHANG*********/

        public void cbxKhachHang_Load()
        {
            cbxMaKH.Items.Clear();
            cbxTenKH.Items.Clear();

            if (conSql.State == ConnectionState.Closed)
            {
                conSql.Open();
            }
            string query = "SELECT MaKH FROM KHACHHANG";
            SqlCommand cmd = new SqlCommand(query, conSql);
            SqlDataReader rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                cbxMaKH.Items.Add(rd["MaKH"].ToString());
            }
            rd.Close();

            query = "SELECT TenKH FROM KHACHHANG";
            cmd = new SqlCommand(query, conSql);
            rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                cbxTenKH.Items.Add(rd["TenKH"].ToString());
            }
            if (conSql.State == ConnectionState.Open)
            {
                conSql.Close();
            }
        }

        /********PHIEUNHAPHANG*********/

        public void cbxPNH_Load()
        {
            cbxMaNCC.Items.Clear();
            cbxMaPN.Items.Clear();
            cbxTenNCC.Items.Clear();
            cbxMaCHNhap.Items.Clear();

            if (conSql.State == ConnectionState.Closed)
            {
                conSql.Open();
            }
            string query = "SELECT MaNCC FROM NHACUNGCAP";
            SqlCommand cmd = new SqlCommand(query, conSql);
            SqlDataReader rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                cbxMaNCC.Items.Add(rd["MaNCC"].ToString());
            }
            rd.Close();

            query = "SELECT MaPN FROM PHIEUNHAPHANG";
            cmd = new SqlCommand(query, conSql);
            rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                cbxMaPN.Items.Add(rd["MaPN"].ToString());
            }
            rd.Close();

            query = "SELECT TenNCC FROM NHACUNGCAP";
            cmd = new SqlCommand(query, conSql);
            rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                cbxTenNCC.Items.Add(rd["TenNCC"].ToString());
            }

            if (conSql.State == ConnectionState.Open)
            {
                conSql.Close();
            }
            cbxMaCHNhap.Items.Add(MaCH());
        }



        /********DONHANG*********/

        public int check_ThanhToan()
        {
            if (rbDaTT.Checked)
            {
                return 1;
            }
            return 0;
        }

        public void cbxDH_Load()
        {
            cbxMaDH.Items.Clear();
            cbxMaKHdh.Items.Clear();
            cbxMaSPdh.Items.Clear();
            cbxMaNVban.Items.Clear();
            cbxMaCHban.Items.Clear();

            if (conSql.State == ConnectionState.Closed)
            {
                conSql.Open();
            }
            string query = "SELECT MaDH FROM DONHANG";
            SqlCommand cmd = new SqlCommand(query, conSql);
            SqlDataReader rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                cbxMaDH.Items.Add(rd["MaDH"].ToString());
            }
            rd.Close();

            query = "SELECT MaKH FROM KHACHHANG";
            cmd = new SqlCommand(query, conSql);
            rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                cbxMaKHdh.Items.Add(rd["MaKH"].ToString());
            }
            rd.Close();

            query = "SELECT MaNV FROM NHANVIEN";
            cmd = new SqlCommand(query, conSql);
            rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                cbxMaNVban.Items.Add(rd["MaNV"].ToString());
            }
            rd.Close();

            query = "SELECT MaSP FROM SANPHAM";
            cmd = new SqlCommand(query, conSql);
            rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                cbxMaSPdh.Items.Add(rd["MaSP"].ToString());
            }
            if (conSql.State == ConnectionState.Open)
            {
                conSql.Close();
            }
            cbxMaCHban.Items.Add(MaCH());

        }

        public void thanhTien_DH()
        {
            int soluongban = int.Parse(txtSoLuongban.Text);
            int giaban = int.Parse(txtGiaBan.Text);
            txtThanhTiendh.Text = (giaban * soluongban).ToString();
        }



        /****************************/

        private void tabCtrl_QuanLy_SelectedIndedChanged(object sender, EventArgs e)
        {
            string query;
            switch (tabCtrl_QuanLy.SelectedIndex)
            {
                case 0:
                    query = "SELECT MaSP AS N'Mã sản phẩm', TenSP AS N'Tên sản phẩm' FROM SANPHAM";
                    dtgList_Load(dtgListSP, query);
                    cbxSanPham_Load();
                    break;
                case 1:

                    query = "SELECT MaNCC AS N'Mã NCC', TenNCC AS N'Nhà cung cấp', SDT AS N'Số điện thoại', DiaChi AS N'Địa chỉ' FROM NHACUNGCAP ";
                    dtgList_Load(dtgListNCC, query);

                    query = loadPNH;
                    dtgList_Load(dtgListPN, query);
                    cbxPNH_Load();
                    guiMa(cbxMaSP, cbxMaSPNhap);

                    break;
                case 2:
                    query = "SELECT MaKH AS N'Mã khách hàng', TenKH AS N'Họ tên khách hàng', SDT AS N'Số điện thoại' FROM KHACHHANG";
                    dtgList_Load(dtgListKH, query);
                    cbxKhachHang_Load();
                    break;
                case 3:

                    dtgList_Load(dtgListDH, loadDH);
                    cbxDH_Load();

                    break;
                case 4:
                    query = "SELECT MaCV AS N'Mã chức vụ', TenCV AS N'Tên chức vụ', MucLuong AS N'Mức lương' FROM CHUCVU ";
                    dtgList_Load(dtgListCV, query);

                    dtgList_Load(dtgListNV, loadNV);
                    cbxNhanVien_Load();
                    break;

            }
        }




        /**************NHANVIEN*****************/
        string loadNV = "SELECT MaNV AS N'Mã nhân viên', TenNV AS N'Tên nhân viên', NamSinh AS N'Năm sinh', " +
                               "NHANVIEN.SDT AS N'Số điện thoại', NHANVIEN.DiaChi AS N'Địa chỉ', NgayNhanViec AS 'Ngày nhận việc', " +
                               "MaCV AS N'Mã chức vụ', TenCV AS N'Tên chức vụ', MaCH AS N'Mã cửa hàng' FROM NHANVIEN " +
                               "JOIN CHUCVU ON MaCV_NV = MaCV " +
                               "JOIN CUAHANG ON MaCH_NV = MaCH";

        private void btThemNV_Click(object sender, EventArgs e)
        {
            if (checkContent(cbxMaNV.Text) && checkContent(txtTenNV.Text) && checkContent(txtSdtNV.Text)
                    && checkContent(txtDiaChiNV.Text) && checkContent(cbxMaCV.Text) && checkContent(cbxMaCHnv.Text))
            {

                try
                {
                    if (conSql.State == ConnectionState.Closed)
                    {
                        conSql.Open();
                    }
                    string namsinh = dtpNgaySinhNV.Text;
                    string ngaynv = dtpNgayNhanViec.Text;
                    string query = "INSERT INTO NHANVIEN VALUES ('" + cbxMaNV.Text + "', N'" + txtTenNV.Text + "', '" + namsinh + "', '" +
                                   txtSdtNV.Text + "', N'" + txtDiaChiNV.Text + "', '" + ngaynv + "', '" +
                                   cbxMaCV.Text + "', '" + cbxMaCHnv.Text + "' )";
                    SqlCommand cmd = new SqlCommand(query, conSql);
                    cmd.ExecuteNonQuery();
                    if (conSql.State == ConnectionState.Open)
                    {
                        conSql.Close();
                    }
                    MessageBox.Show("Them thanh cong!");
                    query = loadNV;
                    dtgList_Load(dtgListNV, query);
                    cbxNhanVien_Load();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ma nhan vien da ton tai. Vui long kiem tra lai!");

                }
            }
            else
            {
                MessageBox.Show("Vui long nhap du thong tin nhan vien!");
            }

        }

        private void btLuuNV_Click(object sender, EventArgs e)
        {
            if (checkContent(cbxMaNV.Text) && checkContent(txtTenNV.Text) && checkContent(txtSdtNV.Text)
                    && checkContent(txtDiaChiNV.Text) && checkContent(cbxMaCV.Text) && checkContent(cbxMaCHnv.Text))
            {

                if (conSql.State == ConnectionState.Closed)
                {
                    conSql.Open();
                }
                string namsinh = dtpNgaySinhNV.Text;
                string ngaynv = dtpNgayNhanViec.Text;
                string query = "UPDATE NHANVIEN SET MaNV = '" + cbxMaNV.Text + "', TenNV = N'" + txtTenNV.Text + "', NamSinh = '" + namsinh + "', SDT = '" +
                                       txtSdtNV.Text + "', DiaChi = N'" + txtDiaChiNV.Text + "', NgayNhanViec = '" + ngaynv + "', MaCV_NV = '" +
                                       cbxMaCV.Text + "', MaCH_NV = '" + cbxMaCHnv.Text + "' WHERE MaNV = '" + cbxMaNV.Text + "' ";
                SqlCommand cmd = new SqlCommand(query, conSql);
                cmd.ExecuteNonQuery();
                MessageBox.Show("Luu thanh cong!");
                if (conSql.State == ConnectionState.Open)
                {
                    conSql.Close();
                }
                query = loadNV;
                dtgList_Load(dtgListNV, query);
                cbxNhanVien_Load();
            }
            else
            {
                MessageBox.Show("Vui long nhap du thong tin nhan vien!");
            }

        }

        private void btXoaNV_Click(object sender, EventArgs e)
        {
            if (conSql.State == ConnectionState.Closed)
            {
                conSql.Open();
            }
            string query = "DELETE FROM NHANVIEN WHERE MaNV = '" + cbxMaNV.Text + "' ";
            SqlCommand cmd = new SqlCommand(query, conSql);
            cmd.ExecuteNonQuery();
            MessageBox.Show("Xoa thanh cong!");
            if (conSql.State == ConnectionState.Open)
            {
                conSql.Close();
            }
            query = loadNV;
            dtgList_Load(dtgListNV, query);
            cbxNhanVien_Load();
        }

        private void btTimNV_Click(object sender, EventArgs e)
        {

            string query = loadNV + " WHERE MaNV = '" + cbxMaNV.Text + "' OR TenNV = N'" + txtTenNV.Text +
                                "' OR MaCV_NV = '" + cbxMaCV.Text + "' OR MaCH_NV = '" + cbxMaCHnv.Text + "' ";
            dtgList_Load(dtgListNV, query);

        }

        private void btAllNV_Click(object sender, EventArgs e)
        {

            dtgList_Load(dtgListNV, loadNV);
        }

        private void dtgListNV_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int i = dtgListNV.CurrentRow.Index;
            cbxMaNV.Text = dtgListNV.Rows[i].Cells[0].Value.ToString();
            txtTenNV.Text = dtgListNV.Rows[i].Cells[1].Value.ToString();
            dtpNgaySinhNV.Text = dtgListNV.Rows[i].Cells[2].Value.ToString();
            txtSdtNV.Text = dtgListNV.Rows[i].Cells[3].Value.ToString();
            txtDiaChiNV.Text = dtgListNV.Rows[i].Cells[4].Value.ToString();
            dtpNgayNhanViec.Text = dtgListNV.Rows[i].Cells[5].Value.ToString();
            cbxMaCV.Text = dtgListNV.Rows[i].Cells[6].Value.ToString();
            cbxTenCV.Text = dtgListNV.Rows[i].Cells[7].Value.ToString();
            cbxMaCHnv.Text = dtgListNV.Rows[i].Cells[8].Value.ToString();

        }


        /**************CHUCVU*****************/

        private void btThemCV_Click(object sender, EventArgs e)
        {
            if (checkContent(cbxMaCV.Text) && checkContent(cbxTenCV.Text) && checkContent(txtMucLuong.Text))
            {
                try
                {

                    if (conSql.State == ConnectionState.Closed)
                    {
                        conSql.Open();
                    }
                    string query = "INSERT INTO CHUCVU VALUES ('" + cbxMaCV.Text + "', N'" + cbxTenCV.Text + "', " + txtMucLuong.Text + " )";
                    SqlCommand cmd = new SqlCommand(query, conSql);
                    cmd.ExecuteNonQuery();
                    if (conSql.State == ConnectionState.Open)
                    {
                        conSql.Close();
                    }
                    MessageBox.Show("Them thanh cong!");
                    query = "SELECT MaCV AS N'Mã chức vụ', TenCV AS N'Tên chức vụ', MucLuong AS N'Mức lương' FROM CHUCVU ";
                    dtgList_Load(dtgListCV, query);
                    cbxNhanVien_Load();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ma chuc vu da ton tai. Vui long kiem tra lai!");
                }
            }
            else
            {
                MessageBox.Show("Vui long nhap day du thong tin chuc vu!");

            }


        }

        private void btLuuCV_Click(object sender, EventArgs e)
        {
            if (checkContent(cbxMaCV.Text) && checkContent(cbxTenCV.Text) && checkContent(txtMucLuong.Text))
            {
                if (conSql.State == ConnectionState.Closed)
                {
                    conSql.Open();
                }
                string query = "UPDATE CHUCVU SET TenCV = N'" + cbxTenCV.Text + "', MucLuong = " + txtMucLuong.Text + " WHERE MaCV = '" + cbxMaCV.Text + "' ";
                SqlCommand cmd = new SqlCommand(query, conSql);
                cmd.ExecuteNonQuery();
                MessageBox.Show("Luu thanh cong!");
                if (conSql.State == ConnectionState.Open)
                {
                    conSql.Close();
                }
                query = "SELECT MaCV AS N'Mã chức vụ', TenCV AS N'Tên chức vụ', MucLuong AS N'Mức lương' FROM CHUCVU ";
                dtgList_Load(dtgListCV, query);
                cbxNhanVien_Load();
            }
            else
            {
                MessageBox.Show("Vui long nhap day du thong tin chuc vu!");
            }
        }

        private void dtgListCV_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int j = dtgListCV.CurrentRow.Index;
            cbxMaCV.Text = dtgListCV.Rows[j].Cells[0].Value.ToString();
            cbxTenCV.Text = dtgListCV.Rows[j].Cells[1].Value.ToString();
            txtMucLuong.Text = dtgListCV.Rows[j].Cells[2].Value.ToString();
        }


        /****************SANPHAM*************/

        private void btThemSP_Click(object sender, EventArgs e)
        {
            if (checkContent(cbxMaSP.Text) && checkContent(cbxTenSP.Text))
            {
                try
                {

                    if (conSql.State == ConnectionState.Closed)
                    {
                        conSql.Open();
                    }
                    string query = "INSERT INTO SANPHAM VALUES ('" + cbxMaSP.Text + "', N'" + cbxTenSP.Text + "' )";
                    SqlCommand cmd = new SqlCommand(query, conSql);
                    cmd.ExecuteNonQuery();
                    if (conSql.State == ConnectionState.Open)
                    {
                        conSql.Close();
                    }
                    MessageBox.Show("Them thanh cong!");
                    query = "SELECT MaSP AS N'Mã sản phẩm', TenSP AS N'Tên sản phẩm' FROM SANPHAM";
                    dtgList_Load(dtgListSP, query);
                    cbxSanPham_Load();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ma san pham da ton tai. Vui long kiem tra lai!");
                }
            }
            else
            {
                MessageBox.Show("Vui long nhap day du thong tin!");
            }
        }

        private void btLuuSP_Click(object sender, EventArgs e)
        {
            if (checkContent(cbxMaSP.Text) && checkContent(cbxTenSP.Text))
            {
                if (conSql.State == ConnectionState.Closed)
                {
                    conSql.Open();
                }
                string query = "UPDATE SANPHAM SET TenSP = N'" + cbxTenSP.Text + "' WHERE MaSP = '" + cbxMaSP.Text + "' ";
                SqlCommand cmd = new SqlCommand(query, conSql);
                cmd.ExecuteNonQuery();
                MessageBox.Show("Luu thanh cong!");
                if (conSql.State == ConnectionState.Open)
                {
                    conSql.Close();
                }
                query = "SELECT MaSP AS N'Mã sản phẩm', TenSP AS N'Tên sản phẩm' FROM SANPHAM";
                dtgList_Load(dtgListSP, query);
            }
            else
            {
                MessageBox.Show("Vui long nhap day du thong tin!");
            }
        }

        private void btXoaSP_Click(object sender, EventArgs e)
        {
            if (conSql.State == ConnectionState.Closed)
            {
                conSql.Open();
            }
            string query = "DELETE FROM SANPHAM WHERE MaSP = '" + cbxMaSP.Text + "' ";
            SqlCommand cmd = new SqlCommand(query, conSql);
            cmd.ExecuteNonQuery();
            MessageBox.Show("Xoa thanh cong!");
            if (conSql.State == ConnectionState.Open)
            {
                conSql.Close();
            }
            query = "SELECT MaSP AS N'Mã sản phẩm', TenSP AS N'Tên sản phẩm' FROM SANPHAM";
            dtgList_Load(dtgListSP, query);
            cbxSanPham_Load();
        }

        private void btTimSP_Click(object sender, EventArgs e)
        {
            string query = "SELECT MaSP AS N'Mã sản phẩm', TenSP AS N'Tên sản phẩm' FROM SANPHAM WHERE MaSP = '" + cbxMaSP.Text + "' OR TenSP = N'" + cbxTenSP.Text + "' ";
            dtgList_Load(dtgListSP, query);
        }

        private void btAllSP_Click(object sender, EventArgs e)
        {
            string query = "SELECT MaSP AS N'Mã sản phẩm', TenSP AS N'Tên sản phẩm' FROM SANPHAM";
            dtgList_Load(dtgListSP, query);
        }

        private void dtgListSP_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int i = dtgListSP.CurrentRow.Index;
            cbxMaSP.Text = dtgListSP.Rows[i].Cells[0].Value.ToString();
            cbxTenSP.Text = dtgListSP.Rows[i].Cells[1].Value.ToString();
        }


        /****************KHACHHANG**************/

        private void btThemKH_Click(object sender, EventArgs e)
        {
            if (checkContent(cbxMaKH.Text) && checkContent(cbxTenKH.Text) && checkContent(txtSdtKH.Text))
            {
                try
                {

                    if (conSql.State == ConnectionState.Closed)
                    {
                        conSql.Open();
                    }
                    string query = "INSERT INTO KHACHHANG VALUES ('" + cbxMaKH.Text + "', N'" + cbxTenKH.Text + "', '" + txtSdtKH.Text + "' )";
                    SqlCommand cmd = new SqlCommand(query, conSql);
                    cmd.ExecuteNonQuery();
                    if (conSql.State == ConnectionState.Open)
                    {
                        conSql.Close();
                    }
                    MessageBox.Show("Them thanh cong!");
                    query = "SELECT MaKH AS N'Mã khách hàng', TenKH AS N'Họ tên khách hàng', SDT AS N'Số điện thoại' FROM KHACHHANG";
                    dtgList_Load(dtgListKH, query);
                    cbxKhachHang_Load();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ma khach hang da ton tai. Vui long kiem tra lai!");
                }
            }
            else
            {
                MessageBox.Show("Vui long nhap day du thong tin!");
            }
        }

        private void btLuuKH_Click(object sender, EventArgs e)
        {
            if (checkContent(cbxMaKH.Text) && checkContent(cbxTenKH.Text) && checkContent(txtSdtKH.Text))
            {
                if (conSql.State == ConnectionState.Closed)
                {
                    conSql.Open();
                }
                string query = "UPDATE KHACHHANG SET TenKH = N'" + cbxTenKH.Text + "', SDT = '" + txtSdtKH.Text + "' WHERE MaKH = '" + cbxMaKH.Text + "' ";
                SqlCommand cmd = new SqlCommand(query, conSql);
                cmd.ExecuteNonQuery();
                MessageBox.Show("Luu thanh cong!");
                if (conSql.State == ConnectionState.Open)
                {
                    conSql.Close();
                }
                query = "SELECT MaKH AS N'Mã khách hàng', TenKH AS N'Họ tên khách hàng', SDT AS N'Số điện thoại' FROM KHACHHANG";
                dtgList_Load(dtgListKH, query);
            }
            else
            {
                MessageBox.Show("Vui long nhap day du thong tin!");
            }
        }

        private void btXoaKH_Click(object sender, EventArgs e)
        {
            if (conSql.State == ConnectionState.Closed)
            {
                conSql.Open();
            }
            string query = "DELETE FROM KHACHHANG WHERE MaKH= '" + cbxMaKH.Text + "'";
            SqlCommand cmd = new SqlCommand(query, conSql);
            cmd.ExecuteNonQuery();
            if (conSql.State == ConnectionState.Open)
            {
                conSql.Close();
            }
            MessageBox.Show("Xoa thanh cong!");
            query = "SELECT MaKH AS N'Mã khách hàng', TenKH AS N'Họ tên khách hàng', SDT AS N'Số điện thoại' FROM KHACHHANG";
            dtgList_Load(dtgListKH, query);
            cbxKhachHang_Load();
        }

        private void btTimKH_Click(object sender, EventArgs e)
        {
            string query = "SELECT MaKH AS N'Mã khách hàng', TenKH AS N'Tên khách hàng', SDT AS N'Số điện thoại' " +
                            "FROM KHACHHANG WHERE MaKH = '" + cbxMaKH.Text + "' OR TenKH = '" + cbxTenKH.Text + "' ";
            dtgList_Load(dtgListKH, query);
        }

        private void btAllKH_Click(object sender, EventArgs e)
        {
            string query = "SELECT MaKH AS N'Mã khách hàng', TenKH AS N'Họ tên khách hàng', SDT AS N'Số điện thoại' FROM KHACHHANG";
            dtgList_Load(dtgListKH, query);
        }

        private void dtgListKH_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int i = dtgListKH.CurrentRow.Index;
            cbxMaKH.Text = dtgListKH.Rows[i].Cells[0].Value.ToString();
            cbxTenKH.Text = dtgListKH.Rows[i].Cells[1].Value.ToString();
            txtSdtKH.Text = dtgListKH.Rows[i].Cells[2].Value.ToString();
        }

        /****************NHACUNGCAP****************/
        private void btThemNCC_Click(object sender, EventArgs e)
        {
            if (checkContent(cbxMaNCC.Text) && checkContent(cbxTenNCC.Text) && checkContent(txtSdtNCC.Text) && checkContent(txtDiaChiNCC.Text))
            {
                try
                {

                    if (conSql.State == ConnectionState.Closed)
                    {
                        conSql.Open();
                    }
                    string query = "INSERT INTO NHACUNGCAP VALUES ('" + cbxMaNCC.Text + "', N'" + cbxTenNCC.Text + "', '" + txtSdtNCC.Text + "', N'" + txtDiaChiNCC.Text + "' )";
                    SqlCommand cmd = new SqlCommand(query, conSql);
                    cmd.ExecuteNonQuery();
                    if (conSql.State == ConnectionState.Open)
                    {
                        conSql.Close();
                    }
                    MessageBox.Show("Them thanh cong!");
                    query = "SELECT MaNCC AS N'Mã NCC', TenNCC AS N'Tên NCC', SDT AS N'Số điện thoại', DiaChi AS N'Địa chỉ' FROM NHACUNGCAP ";
                    dtgList_Load(dtgListNCC, query);
                    cbxPNH_Load();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ma nha cung cap da ton tai. Vui long kiem tra lai!");
                }
            }
            else
            {
                MessageBox.Show("Vui long nhap day du thong tin Nha cung cap!");

            }
        }

        private void btLuuNCC_Click(object sender, EventArgs e)
        {
            if (checkContent(cbxMaNCC.Text) && checkContent(cbxTenNCC.Text) && checkContent(txtSdtNCC.Text) && checkContent(txtDiaChiNCC.Text))
            {
                if (conSql.State == ConnectionState.Closed)
                {
                    conSql.Open();
                }
                string query = "UPDATE NHACUNGCAP SET TenNCC = N'" + cbxTenNCC.Text + "', SDT = '" + txtSdtNCC.Text + "', DiaChi = N'" + txtDiaChiNCC.Text + "' WHERE MaNCC = '" + cbxMaNCC.Text + "' ";
                SqlCommand cmd = new SqlCommand(query, conSql);
                cmd.ExecuteNonQuery();
                MessageBox.Show("Luu thanh cong!");
                if (conSql.State == ConnectionState.Open)
                {
                    conSql.Close();
                }
                query = "SELECT MaNCC AS N'Mã NCC', TenNCC AS N'Tên NCC', SDT AS N'Số điện thoại', DiaChi AS N'Địa chỉ' FROM NHACUNGCAP ";
                dtgList_Load(dtgListNCC, query);
            }
            else
            {
                MessageBox.Show("Vui long nhap day du thong tin Nha cung cap!");
            }
        }

        private void btXoaNCC_Click(object sender, EventArgs e)
        {
            if (conSql.State == ConnectionState.Closed)
            {
                conSql.Open();
            }
            string query = "DELETE FROM NHACUNGCAP WHERE MaNCC = '" + cbxMaNCC.Text + "'";
            SqlCommand cmd = new SqlCommand(query, conSql);
            cmd.ExecuteNonQuery();
            if (conSql.State == ConnectionState.Open)
            {
                conSql.Close();
            }
            MessageBox.Show("Xoa thanh cong!");
            query = "SELECT MaNCC AS N'Mã NCC', TenNCC AS N'Tên NCC', SDT AS N'Số điện thoại', DiaChi AS N'Địa chỉ' FROM NHACUNGCAP ";
            dtgList_Load(dtgListNCC, query);
            cbxPNH_Load();
        }

        private void btTimNCC_Click(object sender, EventArgs e)
        {
            string query = "SELECT MaNCC AS N'Mã NCC', TenNCC AS N'Tên NCC', SDT AS N'Số điện thoại', DiaChi AS N'Địa chỉ' FROM NHACUNGCAP " +
                "WHERE MaNCC = '" + cbxMaNCC.Text + "' OR TenNCC = N'" + cbxTenNCC.Text + "' ";
            dtgList_Load(dtgListNCC, query);
        }

        private void dtgListNCC_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int i = dtgListNCC.CurrentRow.Index;
            cbxMaNCC.Text = dtgListNCC.Rows[i].Cells[0].Value.ToString();
            cbxTenNCC.Text = dtgListNCC.Rows[i].Cells[1].Value.ToString();
            txtSdtNCC.Text = dtgListNCC.Rows[i].Cells[2].Value.ToString();
            txtDiaChiNCC.Text = dtgListNCC.Rows[i].Cells[3].Value.ToString();
        }

        /***************PHIEUNHAPHANG*************/

        string loadPNH = "SELECT MaPN AS N'Mã phiếu nhập', NgayNhap AS N'Ngày nhập',MaSP AS N'Mã sản phẩm', TenSP AS N'Sản phẩm', " +
                            "GiaNhap AS N'Giá nhập', SoLuong AS N'Số lượng', ThanhTien AS N'Thành tiền', NSX, HSD, " +
                            "MaNCC AS N'Mã NCC',TenNCC AS N'Nhà cung cấp', MaCH_PN AS 'Mã cửa hàng' FROM PHIEUNHAPHANG " +
                            "JOIN NHACUNGCAP ON MaNCC_PN = MaNCC " +
                            "JOIN CHITIET_PHIEUNHAP ON MaPN = MaPN_CTPN " +
                            "JOIN CUAHANG ON MaCH_PN = MaCH " +
                            "JOIN SANPHAM ON MaSP_CTPN = MaSP ";
        private void btThemPN_Click(object sender, EventArgs e)
        {
            if (checkContent(cbxMaPN.Text) && checkContent(cbxMaNCC.Text) && checkContent(cbxMaSPNhap.Text)
                && checkContent(txtGiaNhap.Text) && checkContent(txtSoLuongNhap.Text) && checkContent(cbxMaCHNhap.Text))
            {
                try
                {

                    if (conSql.State == ConnectionState.Closed)
                    {
                        conSql.Open();
                    }

                    string query = "INSERT INTO PHIEUNHAPHANG VALUES ('" + cbxMaPN.Text + "', '" + dtpNgayNhap.Text + "', '" + cbxMaNCC.Text + "', '" + cbxMaCHNhap.Text + "' )";
                    SqlCommand cmd = new SqlCommand(query, conSql);
                    cmd.ExecuteNonQuery();

                    query = "INSERT INTO CHITIET_PHIEUNHAP VALUES ('" + cbxMaPN.Text + "', '" + cbxMaSPNhap.Text + "', " +
                            txtGiaNhap.Text + ", " + txtSoLuongNhap.Text + ", " + txtThanhTienNhap.Text + ", '" +
                            dtpNSX.Text + "', '" + dtpHSD.Text + "' )";
                    cmd = new SqlCommand(query, conSql);
                    cmd.ExecuteNonQuery();

                    MessageBox.Show("Them thanh cong!");
                    query = loadPNH;
                    dtgList_Load(dtgListPN, query);
                    cbxPNH_Load();

                    if (conSql.State == ConnectionState.Open)
                    {
                        conSql.Close();
                    }
                }
                catch (Exception ex)
                {
                    if (conSql.State == ConnectionState.Closed)
                    {
                        conSql.Open();
                    }

                    string query = "INSERT INTO CHITIET_PHIEUNHAP VALUES ('" + cbxMaPN.Text + "', '" + cbxMaSPNhap.Text + "', " +
                            txtGiaNhap.Text + ", " + txtSoLuongNhap.Text + ", " + txtThanhTienNhap.Text + ", '" +
                            dtpNSX.Text + "', '" + dtpHSD.Text + "' )";
                    SqlCommand cmd = new SqlCommand(query, conSql);
                    cmd = new SqlCommand(query, conSql);
                    cmd.ExecuteNonQuery();

                    MessageBox.Show("Them thanh cong!");
                    query = loadPNH;
                    dtgList_Load(dtgListPN, query);
                    cbxPNH_Load();

                    if (conSql.State == ConnectionState.Open)
                    {
                        conSql.Close();
                    }
                }
            }
            else
            {
                MessageBox.Show("Vui long nhap day du thong tin Phieu nhap hang!");

            }
        }

        private void btLuuPN_Click(object sender, EventArgs e)
        {
            if (checkContent(cbxMaPN.Text) && checkContent(cbxMaNCC.Text) && checkContent(cbxMaSPNhap.Text)
                && checkContent(txtGiaNhap.Text) && checkContent(txtSoLuongNhap.Text) && checkContent(cbxMaCHNhap.Text))
            {
                if (conSql.State == ConnectionState.Closed)
                {
                    conSql.Open();
                }
                string query = "UPDATE PHIEUNHAPHANG SET NgayNhap = '" + dtpNgayNhap.Text + "', MaNCC_PN = '" + cbxMaNCC.Text + "', MaCH_PN = '" + cbxMaCHNhap.Text + "' WHERE MaPN = '" + cbxMaPN.Text + "' ";
                SqlCommand cmd = new SqlCommand(query, conSql);
                cmd.ExecuteNonQuery();

                query = "UPDATE CHITIET_PHIEUNHAP SET Gianhap = " +
                            txtGiaNhap.Text + ", SoLuong = " + txtSoLuongNhap.Text + ", ThanhTien = " + txtThanhTienNhap.Text + ", NSX = '" +
                            dtpNSX.Text + "', HSD = '" + dtpHSD.Text + "' WHERE MaPN_CTPN = '" + cbxMaPN.Text + "' AND MaSP_CTPN = '" + cbxMaSPNhap.Text + "' ";
                cmd = new SqlCommand(query, conSql);
                cmd.ExecuteNonQuery();

                MessageBox.Show("Luu thanh cong!");
                if (conSql.State == ConnectionState.Open)
                {
                    conSql.Close();
                }
                query = loadPNH;
                dtgList_Load(dtgListPN, query);
            }
            else
            {
                MessageBox.Show("Vui long nhap day du thong tin Phieu nhap hang!");
            }
        }

        private void btXoaPN_Click(object sender, EventArgs e)
        {
            if (conSql.State == ConnectionState.Closed)
            {
                conSql.Open();
            }
            string query = "DELETE FROM CHITIET_PHIEUNHAP WHERE MaPN_CTPN = '" + cbxMaPN.Text + "'";
            SqlCommand cmd = new SqlCommand(query, conSql);
            cmd.ExecuteNonQuery();

            query = "DELETE FROM PHIEUNHAPHANG WHERE MaPN = '" + cbxMaPN.Text + "'";
            cmd = new SqlCommand(query, conSql);
            cmd.ExecuteNonQuery();

            if (conSql.State == ConnectionState.Open)
            {
                conSql.Close();
            }
            MessageBox.Show("Xoa thanh cong!");
            query = loadPNH;
            dtgList_Load(dtgListPN, query);
            cbxPNH_Load();
        }

        private void btTimPN_Click(object sender, EventArgs e)
        {
            string query = loadPNH +
                            " WHERE MaPN = '" + cbxMaPN.Text + "' OR MaSP_CTPN = '" + cbxMaSP.Text + "' OR " +
                            "MaNCC_PN = '" + cbxMaNCC.Text + "' OR MaCH_PN = '" + cbxMaCHNhap.Text + "' ";
            dtgList_Load(dtgListPN, query);
        }

        private void btAllPN_Click(object sender, EventArgs e)
        {
            string query;
            query = "SELECT MaNCC AS N'Mã NCC', TenNCC AS N'Tên NCC', SDT AS N'Số điện thoại', DiaChi AS N'Địa chỉ' FROM NHACUNGCAP ";
            dtgList_Load(dtgListNCC, query);

            query = loadPNH;
            dtgList_Load(dtgListPN, query);
        }

        private void dtgListPN_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int i = dtgListPN.CurrentRow.Index;
            cbxMaPN.Text = dtgListPN.Rows[i].Cells[0].Value.ToString();
            dtpNgayNhap.Text = dtgListPN.Rows[i].Cells[1].Value.ToString();
            cbxMaSPNhap.Text = dtgListPN.Rows[i].Cells[2].Value.ToString();
            txtGiaNhap.Text = dtgListPN.Rows[i].Cells[4].Value.ToString();
            txtSoLuongNhap.Text = dtgListPN.Rows[i].Cells[5].Value.ToString(); ;
            txtThanhTienNhap.Text = dtgListPN.Rows[i].Cells[6].Value.ToString();
            dtpNSX.Text = dtgListPN.Rows[i].Cells[7].Value.ToString();
            dtpHSD.Text = dtgListPN.Rows[i].Cells[8].Value.ToString();
            cbxMaNCC.Text = dtgListPN.Rows[i].Cells[9].Value.ToString();
            cbxTenNCC.Text = dtgListPN.Rows[i].Cells[10].Value.ToString();
            cbxMaCHNhap.Text = dtgListPN.Rows[i].Cells[11].Value.ToString();
        }

        private void txtSoLuongNhap_TextChanged(object sender, EventArgs e)
        {
            if (checkContent(txtGiaNhap.Text) && checkContent(txtSoLuongNhap.Text))
            {
                thanhTien_PN();
            }

        }

        private void txtGiaNhap_TextChanged(object sender, EventArgs e)
        {
            if (checkContent(txtGiaNhap.Text) && checkContent(txtSoLuongNhap.Text))
            {
                thanhTien_PN();
            }
        }


        /*****************DONHANG*************/

        string loadDH = "SELECT MaDH AS N'Mã đơn hàng', NgayDatHang AS N'Ngày đặt hàng',MaKH AS N'Mã khách hàng', TenKH AS N'Khách hàng', MaSP AS N'Mã sản phẩm', " +
                        "TenSP AS N'Sản phẩm', GiaBan AS N'Giá bán', SoLuong AS N'Số lượng', ThanhTien AS N'Thành tiền', " +
                        "ThanhToan AS N'Thanh toán', DiaChi AS N'Địa chỉ giao hàng', MaNV_DH AS N'Mã nhân viên', MaCH_DH AS N'Mã cửa hàng' " +
                        "FROM DONHANG JOIN CHITIET_DONHANG ON MaDH = MaDH_CTDH " +
                        "JOIN KHACHHANG ON MaKH_DH = MaKH " +
                        "JOIN SANPHAM ON MaSP_CTDH = MaSP ";
        private void btThemDH_Click(object sender, EventArgs e)
        {
            if (checkContent(cbxMaDH.Text) && checkContent(cbxMaKHdh.Text) && checkContent(cbxMaSPdh.Text)
                && checkContent(txtGiaBan.Text) && checkContent(txtSoLuongban.Text) && checkContent(cbxMaCHban.Text)
                && checkContent(cbxMaNVban.Text) && checkContent(txtDiaChiGH.Text))
            {
                try
                {

                    if (conSql.State == ConnectionState.Closed)
                    {
                        conSql.Open();
                    }

                    string query = "INSERT INTO DONHANG VALUES ('" + cbxMaDH.Text + "', '" + dtpTGDatHang.Text + "', N'" +
                                    txtDiaChiGH.Text + "', '" + cbxMaKHdh.Text + "', '" + cbxMaNVban.Text + "', '" +
                                    cbxMaCHban.Text + "' )";
                    SqlCommand cmd = new SqlCommand(query, conSql);
                    cmd.ExecuteNonQuery();

                    query = "INSERT INTO CHITIET_DONHANG VALUES ('" + cbxMaDH.Text + "', '" + cbxMaSPdh.Text + "', " +
                            txtGiaBan.Text + ", " + txtSoLuongban.Text + ", " + txtThanhTiendh.Text + ", " +
                            check_ThanhToan().ToString() + " )";
                    cmd = new SqlCommand(query, conSql);
                    cmd.ExecuteNonQuery();

                    MessageBox.Show("Them thanh cong!");
                    query = loadDH;
                    dtgList_Load(dtgListDH, query);
                    cbxDH_Load();

                    if (conSql.State == ConnectionState.Open)
                    {
                        conSql.Close();
                    }
                }
                catch (Exception ex)
                {
                    try
                    {

                        if (conSql.State == ConnectionState.Closed)
                        {
                            conSql.Open();
                        }

                        string query = "INSERT INTO CHITIET_DONHANG VALUES ('" + cbxMaDH.Text + "', '" + cbxMaSPdh.Text + "', " +
                                txtGiaBan.Text + ", " + txtSoLuongban.Text + ", " + txtThanhTiendh.Text + ", " +
                                check_ThanhToan().ToString() + " )";
                        SqlCommand cmd = new SqlCommand(query, conSql);
                        cmd.ExecuteNonQuery();

                        MessageBox.Show("Them thanh cong!");
                        query = loadDH;
                        dtgList_Load(dtgListDH, query);
                        cbxDH_Load();

                        if (conSql.State == ConnectionState.Open)
                        {
                            conSql.Close();
                        }
                    }
                    catch (Exception ex1)
                    {
                        MessageBox.Show("Ma san pham cua Don hang nay da ton tai. Vui long kiem tra lai!");
                    }
                }
            }
            else
            {
                MessageBox.Show("Vui long nhap day du thong tin Don hang!");

            }

        }

        private void btLuuDH_Click(object sender, EventArgs e)
        {

            if (checkContent(cbxMaDH.Text) && checkContent(cbxMaKHdh.Text) && checkContent(cbxMaSPdh.Text)
                && checkContent(txtGiaBan.Text) && checkContent(txtSoLuongban.Text) && checkContent(cbxMaCHban.Text)
                && checkContent(cbxMaNVban.Text) && checkContent(txtDiaChiGH.Text))
            {
                if (conSql.State == ConnectionState.Closed)
                {
                    conSql.Open();
                }
                string query = "UPDATE DONHANG SET NgayDatHang = '" + dtpTGDatHang.Text + "',  DiaChi = N'" +
                                    txtDiaChiGH.Text + "', MaKH_DH = '" + cbxMaKHdh.Text + "', MaNV_DH = '" + cbxMaNVban.Text + "', MaCH_DH = '" +
                                    cbxMaCHban.Text + "' WHERE MaDH = '" + cbxMaDH.Text + "' ";
                SqlCommand cmd = new SqlCommand(query, conSql);
                cmd.ExecuteNonQuery();

                query = "UPDATE CHITIET_DONHANG SET GiaBan = " + txtGiaBan.Text + ", SoLuong = " + txtSoLuongban.Text + ", ThanhTien = "
                        + txtThanhTiendh.Text + ", ThanhToan = " + check_ThanhToan().ToString()
                        + " WHERE MaDH_CTDH = '" + cbxMaDH.Text + "' AND MaSP_CTDH = '" + cbxMaSPdh.Text + "' ";
                cmd = new SqlCommand(query, conSql);
                cmd.ExecuteNonQuery();

                MessageBox.Show("Luu thanh cong!");
                if (conSql.State == ConnectionState.Open)
                {
                    conSql.Close();
                }
                query = loadDH;
                dtgList_Load(dtgListDH, query);
            }
            else
            {
                MessageBox.Show("Vui long nhap day du thong tin Don hang!");
            }
        }

        private void btXoaDH_Click(object sender, EventArgs e)
        {
            if (conSql.State == ConnectionState.Closed)
            {
                conSql.Open();
            }
            string query = "DELETE FROM CHITIET_DONHANG WHERE MaDH_CTDH = '" + cbxMaDH.Text + "' AND MaSP_CTDH = '" + cbxMaSPdh.Text + "' ";
            SqlCommand cmd = new SqlCommand(query, conSql);
            cmd.ExecuteNonQuery();

            query = "SELECT * FROM CHITIET_DONHANG ";
            cmd = new SqlCommand(query, conSql);
            SqlDataReader rd = cmd.ExecuteReader();
            if (!rd.Read())
            {
                query = "DELETE FROM DONHANG WHERE MaDH = '" + cbxMaDH.Text + "'";
                cmd = new SqlCommand(query, conSql);
                cmd.ExecuteNonQuery();
            }
            rd.Close();

            if (conSql.State == ConnectionState.Open)
            {
                conSql.Close();
            }
            MessageBox.Show("Xoa thanh cong!");
            query = loadDH;
            dtgList_Load(dtgListDH, query);
            cbxDH_Load();
        }

        private void btTimDH_Click(object sender, EventArgs e)
        {
            string query = loadDH + "WHERE MaDH = '" + cbxMaDH.Text + "' OR MaSP_CTDH = '" + cbxMaSPdh.Text + "' OR MaKH_DH = '" +
                            cbxMaKHdh.Text + "' OR MaCH_DH = '" + cbxMaCHban.Text + "' OR MaNV_DH = '" + cbxMaNVban.Text + "' ";
            dtgList_Load(dtgListDH, query);
        }

        private void btAllListDH_Click(object sender, EventArgs e)
        {
            dtgList_Load(dtgListDH, loadDH);
        }

        private void dtgListDH_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int i = dtgListDH.CurrentRow.Index;
            int count = 0;
            cbxMaDH.Text = dtgListDH.Rows[i].Cells[count++].Value.ToString();
            dtpTGDatHang.Text = dtgListDH.Rows[i].Cells[count++].Value.ToString();
            cbxMaKHdh.Text = dtgListDH.Rows[i].Cells[count++].Value.ToString();
            count++;
            cbxMaSPdh.Text = dtgListDH.Rows[i].Cells[count++].Value.ToString();
            count++;
            txtGiaBan.Text = dtgListDH.Rows[i].Cells[count++].Value.ToString();
            txtSoLuongban.Text = dtgListDH.Rows[i].Cells[count++].Value.ToString();
            txtThanhTiendh.Text = dtgListDH.Rows[i].Cells[count++].Value.ToString();
           
            txtDiaChiGH.Text = dtgListDH.Rows[i].Cells[count++].Value.ToString();
            cbxMaNVban.Text = dtgListDH.Rows[i].Cells[count++].Value.ToString(); ;
            cbxMaCHban.Text = dtgListDH.Rows[i].Cells[count++].Value.ToString();
        }

        private void txtGiaBan_TextChanged(object sender, EventArgs e)
        {
            if (checkContent(txtGiaBan.Text) && checkContent(txtSoLuongban.Text))
            {
                thanhTien_DH();
            }
        }

        private void txtSoLuongban_TextChanged(object sender, EventArgs e)
        {
            if (checkContent(txtGiaBan.Text) && checkContent(txtSoLuongban.Text))
            {
                thanhTien_DH();
            }
        }

        /************BUTTON TAO MOI***************/

        private void btTaoMoiSP_Click(object sender, EventArgs e)
        {
            cbxMaSP.Text = "";
            cbxTenSP.Text = "";
        }

        private void btTaoMoiPN_Click(object sender, EventArgs e)
        {
            cbxMaPN.Text = "";
            cbxMaNCC.Text = "";
            cbxMaSPNhap.Text = "";
            cbxMaCHNhap.Text = "";
            txtDiaChiNCC.Text = "";
            txtSdtNCC.Text = "";
            cbxTenNCC.Text = "";
            dtpNgayNhap.Text = "";
            txtSoLuongNhap.Text = "";
            txtGiaNhap.Text = "";
            txtThanhTienNhap.Text = "";
            dtpNSX.Text = "";
            dtpHSD.Text = "";
        }

        private void btTaoMoiKH_Click(object sender, EventArgs e)
        {
            cbxMaKH.Text = "";
            cbxTenKH.Text = "";
            txtSdtKH.Text = "";
        }

        private void btTaoMoiDH_Click(object sender, EventArgs e)
        {
            cbxMaDH.Text = "";
            cbxMaNVban.Text = "";
            cbxMaCHban.Text = "";
            cbxMaSPdh.Text = "";
            cbxMaKHdh.Text = "";
            txtSoLuongban.Text = "";
            txtGiaBan.Text = "";
            txtThanhTiendh.Text = "";
            txtDiaChiGH.Text = "";
            dtpTGDatHang.Text = "";
            rbChuaTT.Refresh();
            rbDaTT.Refresh();
        }

        private void btTaoMoiNV_Click(object sender, EventArgs e)
        {
            cbxMaCV.Text = "";
            cbxTenCV.Text = "";
            cbxMaNV.Text = "";
            cbxMaCHnv.Text = "";
            txtDiaChiNV.Text = "";
            txtSdtNV.Text = "";
            txtTenNV.Text = "";
            dtpNgayNhanViec.Text = "";
            dtpNgaySinhNV.Text = "";
            txtMucLuong.Text = "";
        }
    }
}
