using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace QLyCHBanDoAnVat
{
    public partial class Form1 : Form
    {
        public SqlConnection sqlconn;

        public Form1()
        {
            InitializeComponent();
        }

        /*************************/

        private bool checkTxt(TextBox txt)
        {
            if (txt.Text != "")
            {
                return true;
            }
            return false;
        }

        public bool check_TaiKhoan(SqlConnection con)
        {
            string query = "SELECT SDT FROM NHANVIEN";
            SqlCommand cmd = new SqlCommand(query, con);
            SqlDataReader rd = cmd.ExecuteReader();
            while(rd.Read())
            {
                if (rd["SDT"].ToString() == txtLoginName.Text)
                {
                    return true;
                }
            }
            return false;
        }

        /************************/


        private void Form1_Load(object sender, EventArgs e)
        {
            btDangNhap.Enabled = false;
        }

        private void txtLoginName_TextChanged(object sender, EventArgs e)
        {
            if (checkTxt(txtLoginName) && checkTxt(txtPassword))
            {
                btDangNhap.Enabled = true;
            }
            else
            {
                btDangNhap.Enabled = false;
            }
        }

        private void txtPassword_TextChanged(object sender, EventArgs e)
        {
            if (checkTxt(txtLoginName) && checkTxt(txtPassword))
            {
                btDangNhap.Enabled = true;
            }
            else
            {
                btDangNhap.Enabled = false;
            }
        }
        private void txtLoginName_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar);
        }

        private void btDangNhap_Click(object sender, EventArgs e)
        {
            string conString = "Data Source=TUONGVIB09C; Initial Catalog=CuaHangBanDoAnVat_SQL; User ID='" + txtLoginName.Text + "' ; Password= '" + txtPassword.Text + "';";
            sqlconn = new SqlConnection(conString);
            try
            {
                if (sqlconn.State == ConnectionState.Closed)
                {
                    sqlconn.Open();
                }
                MessageBox.Show("Dang nhap thanh cong!");

                if (check_TaiKhoan(sqlconn))
                {
                    Form_NguoiQuanLy quanly = new Form_NguoiQuanLy(conString, txtLoginName.Text);
                    quanly.Show();
                }
                else { 
                Form_TrangChu trangChu = new Form_TrangChu(conString);
                trangChu.Show();
                }
                if (sqlconn.State == ConnectionState.Open)
                {
                    sqlconn.Close();
                }
            }
            catch(Exception ex) 
            {
                MessageBox.Show("Ten dang nhap hoac mat khau khong dung. Vui long kiem tra lai!");
            }
        }
    }
}