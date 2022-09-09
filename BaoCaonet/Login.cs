using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BaoCaonet
{
    public partial class Login : Form
    {
        public Login()
        {
            InitializeComponent();
        }

       

        //
       
        private void btnLogin_Click_1(object sender, EventArgs e)
        {
            try
            {
                //B1 Khởi tạo kết nối
                string connect = @"Data Source=DESKTOP-ELOR9UD\SQLEXPRESS;Initial Catalog=QuanLyChungCu2;Integrated Security=True";
                SqlConnection conn = new SqlConnection(connect);
                //B2 Khởi tạo kết nối
                conn.Open();
                //B3 Tạo truy vấn
                string sql = "select VaiTro from TAIKHOAN where TenTaiKhoan='" + txtUsername.Text + "' and MatKhau='" + txtPassword.Text + "'";
                //B4 Thực thi truy vấn
                string sql0 = "select COUNT(*) from TAIKHOAN where TenTaiKhoan='" + txtUsername.Text + "'";
                SqlCommand cmd = new SqlCommand(sql, conn);
                SqlCommand cmd0 = new SqlCommand(sql0, conn);

                int check = (int)cmd0.ExecuteScalar();
                if (check == 1)
                {
                    var VaiTro = (int)cmd.ExecuteScalar();
                    if (VaiTro == 1)
                    {
                        MessageBox.Show("Đăng Nhập Admin");
                        Admin x = new Admin();
                        x.Show();
                        this.SetVisibleCore(false);
                    }
                    else if (VaiTro == 0)
                    {
                        MessageBox.Show("Đăng Nhập Nhân Viên");
                        NhanVien.NhanVien x = new NhanVien.NhanVien();
                        x.Show();
                        this.SetVisibleCore(false);
                    }
                    else
                    {
                        MessageBox.Show("Tài khoản hoặc mật khẩu sai");
                    }
                }
                else
                {
                    MessageBox.Show("Tài khoản hoặc mật khẩu sai");
                }

                conn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi:" + ex.Message);
            }
        }
        //
    }
}
