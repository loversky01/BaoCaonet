using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
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

        //Kiem tra login
        public bool checkObject()
        {
            if (string.IsNullOrWhiteSpace(txtUsername.Text))
            {
                MessageBox.Show("Bạn chưa nhập tài khoản", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtUsername.Focus();
                return false;
            }
            if (string.IsNullOrWhiteSpace(txtPassword.Text))
            {
                MessageBox.Show("Bạn chưa nhập mật khẩu", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtPassword.Focus();
                return false;
            }
            return true;
        }

        //
       
        private void btnLogin_Click_1(object sender, EventArgs e)
        {
            if (checkObject())
            {
                string user = txtUsername.Text;
                string pass = txtPassword.Text;
                string sql = "select * from TAIKHOAN where TenTaiKhoan='" + user + "' and MatKhau='" + pass + "' and VaiTro='True' ";
                DataTable mytable = Ketnoi.SelectDB(sql);
                if (mytable.Rows.Count > 0)
                {
                    Admin f = new Admin();
                    f.Show();
                    this.Hide();
                }
                else
                {
                    MessageBox.Show("Tên đăng nhập hoặc mật khẩu không chính xác!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
            }
        }
        //
    }
}
