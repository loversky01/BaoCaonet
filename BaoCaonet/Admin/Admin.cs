using BaoCaonet.Entity;
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
    public partial class Admin : Form
    {
        public Admin()
        {
            InitializeComponent();
        }

        public void showKhuCanHo()
        {
            List<KhuCanHoDAL> list = new List<KhuCanHoDAL>();
            SqlConnection conn = new SqlConnection(@"Data Source=DESKTOP-ELOR9UD\SQLEXPRESS;Initial Catalog=QuanLyChungCu2;Integrated Security=True");
            conn.Open();
            string query = "select * from KHUCANHO";
            SqlCommand cmd = new SqlCommand(query, conn);
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                KhuCanHoDAL a = new KhuCanHoDAL();
                a.MaKhu = dr["MaKhu"].ToString();
                a.TenKhu = dr["TenKhu"].ToString();
                a.SoTang = (int)dr["SoTang"];
                a.SoCanTT = (int)dr["SoCanTT"];
                a.DiaChi = dr["DiaChi"].ToString();
                list.Add(a);
            }
            conn.Close();
            dgvKhucanho.DataSource = list;
        }

        //Show thong tin mua ban
        public void showThongTinMuaBan()
        {
            List<ThongTinMuaBanDAL> list = new List<ThongTinMuaBanDAL>();
            SqlConnection conn = new SqlConnection(@"Data Source=DESKTOP-ELOR9UD\SQLEXPRESS;Initial Catalog=QuanLyChungCu2;Integrated Security=True");
            conn.Open();
            string query = "select distinct CUDAN.TenCuDan,HOPDONG.MaHopDong,CUDAN.MaCuDan,CANHO.MaCanHo,CANHO.DienTich,HOPDONG.DiaChiKH,HOPDONG.NgayGiaoDich ,CANHO.Gia from CUDAN, CANHO, HOPDONG where CUDAN.MaCuDan = HOPDONG.MaCuDan and CANHO.MaCanHo = HOPDONG.MaCanHo";
            SqlCommand cmd = new SqlCommand(query, conn);
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                ThongTinMuaBanDAL a = new ThongTinMuaBanDAL();
                a.MaCuDan = dr["MaCuDan"].ToString();
                a.TenCuDan = dr["TenCuDan"].ToString();
                a.MaHopDong = dr["MaHopDong"].ToString();
                a.MaCanHo = dr["MaCanHo"].ToString();
                a.DiaChiKH = dr["DiaChiKH"].ToString();
                a.Gia = (long)dr["Gia"];
                a.NgayGiaoDich = dr["NgayGiaoDich"].ToString();
                list.Add(a);
            }
            conn.Close();
            dgvThongTinMuaBan.DataSource = list;
        }


        //Show thong tin cu dan
        public void showThongTinCuDan()
        {
            List<ThongTinCuDanDAL> list = new List<ThongTinCuDanDAL>();
            SqlConnection conn = new SqlConnection(@"Data Source=DESKTOP-ELOR9UD\SQLEXPRESS;Initial Catalog=QuanLyChungCu2;Integrated Security=True");
            conn.Open();
            string query = "select * from CUDAN";
            SqlCommand cmd = new SqlCommand(query, conn);
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                ThongTinCuDanDAL a = new ThongTinCuDanDAL();
                a.MaCuDan = dr["MaCuDan"].ToString();
                a.TenCuDan = dr["TenCuDan"].ToString();
                a.NgaySinh = dr["NgaySinh"].ToString();
                if (dr["GioiTinh"].ToString() == "True")
                {
                    a.GioiTinh = "Nam";
                }
                else
                {
                    a.GioiTinh = "Nữ";
                }
                a.SoDT = dr["SoDT"].ToString();
                a.SoCMT = dr["SoCMT"].ToString();
                a.QueQuan = dr["QueQuan"].ToString();
                list.Add(a);
            }
            conn.Close();
            dgvThongTinCuDan.DataSource = list;
        }
        //Show thông tin tài khoản 
        public void showTK()
        {
            List<TaiKhoanDAL> list = new List<TaiKhoanDAL>();
            SqlConnection conn = new SqlConnection(@"Data Source=DESKTOP-ELOR9UD\SQLEXPRESS;Initial Catalog=QuanLyChungCu2;Integrated Security=True");
            conn.Open();
            string query = "select * from TAIKHOAN";
            SqlCommand cmd = new SqlCommand(query, conn);
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                TaiKhoanDAL a = new TaiKhoanDAL();
                a.TenTaiKhoan = dr["TenTaiKhoan"].ToString();
                a.MatKhau = dr["MatKhau"].ToString();
                a.VaiTro = dr["Vaitro"].ToString();
                list.Add(a);
            }
            conn.Close();
            dgvTK.DataSource = list;
        }
        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void Admin_Load(object sender, EventArgs e)
        {
            showKhuCanHo();
            showCanHo();
            showThongTinMuaBan();
            showThongTinCuDan();
            showTK();
        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void btnThemKhuCH_Click(object sender, EventArgs e)
        {
            SqlConnection conn = new SqlConnection(@"Data Source=DESKTOP-ELOR9UD\SQLEXPRESS;Initial Catalog=QuanLyChungCu2;Integrated Security=True");
            conn.Open();
            try
            {
                string query = "INSERT INTO KHUCANHO (MaKhu, TenKhu, SoTang,SoCanTT,DiaChi)VALUES ('" + txtMaKhu.Text + "',N'" + txtTenKhu.Text + "','" + txtSoTang.Text + "','" + txtSoCan.Text + "',N'" + txtDiaChi.Text + "')";
                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.ExecuteNonQuery();
                MessageBox.Show("Thêm khu căn hộ thành công");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi:" + ex.Message);
            }
            conn.Close();
            showKhuCanHo();
        }

        private void btnTimKiem_Click(object sender, EventArgs e)
        {
            List<KhuCanHoDAL> list = new List<KhuCanHoDAL>();
            SqlConnection conn = new SqlConnection(@"Data Source=DESKTOP-ELOR9UD\SQLEXPRESS;Initial Catalog=QuanLyChungCu2;Integrated Security=True");
            conn.Open();
            string query = "select * from KHUCANHO where MaKhu like '%" + txtTimkiemKCH.Text + "%' or TenKhu like'%" + txtTimkiemKCH.Text + "%' or SoTang like '%" + txtTimkiemKCH.Text + "%' or SoCanTT like'%" + txtTimkiemKCH.Text + "%'or MaKhu like'%" + txtTimkiemKCH.Text + "%'";
            SqlCommand cmd = new SqlCommand(query, conn);
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                KhuCanHoDAL a = new KhuCanHoDAL();
                a.MaKhu = dr["MaKhu"].ToString();
                a.TenKhu = dr["TenKhu"].ToString();
                a.SoTang = (int)dr["SoTang"];
                a.SoCanTT = (int)dr["SoCanTT"];
                a.DiaChi = dr["DiaChi"].ToString();
                list.Add(a);
            }
            conn.Close();
            dgvKhucanho.DataSource = list;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void dgvKhucanho_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            int idx = e.RowIndex;
            txtMaKhu.Enabled = false;
            txtMaKhu.Text = dgvKhucanho.Rows[idx].Cells["MaKhu"].Value.ToString();
            txtTenKhu.Text = dgvKhucanho.Rows[idx].Cells["TenKhu"].Value.ToString();
            txtSoTang.Text = dgvKhucanho.Rows[idx].Cells["SoTang"].Value.ToString();
            txtSoCan.Text = dgvKhucanho.Rows[idx].Cells["SoCanTT"].Value.ToString();
            txtDiaChi.Text = dgvKhucanho.Rows[idx].Cells["DiaChi"].Value.ToString();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            txtMaKhu.Enabled = true;
            txtMaKhu.Clear();
            txtMaKhu.Clear();
            txtTenKhu.Clear();
            txtSoTang.Clear();
            txtSoCan.Clear();
            txtDiaChi.Clear();

            showKhuCanHo();
        }

        private void dgvCanHo_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            int idx = e.RowIndex;
            txtMCH.Enabled = false;
            txtMK.Enabled = false;
            txtMCH.Text = dgvCanHo.Rows[idx].Cells["MaCanHo"].Value.ToString();
            txtDT.Text = dgvCanHo.Rows[idx].Cells["DienTich"].Value.ToString();
            txtGia.Text = dgvCanHo.Rows[idx].Cells["Gia"].Value.ToString();
            txtSP.Text = dgvCanHo.Rows[idx].Cells["SoPhong"].Value.ToString();
            txtMCD.Text = dgvCanHo.Rows[idx].Cells["MaCuDan"].Value.ToString();
            txtMK.Text = dgvCanHo.Rows[idx].Cells["MaKhu"].Value.ToString();
        }

        //QuanLyCanHo
        public void showCanHo()
        {
            List<CanHoDAL> list = new List<CanHoDAL>();
            SqlConnection conn = new SqlConnection(@"Data Source=DESKTOP-ELOR9UD\SQLEXPRESS;Initial Catalog=QuanLyChungCu2;Integrated Security=True");
            conn.Open();
            string query = "select * from CANHO";
            SqlCommand cmd = new SqlCommand(query, conn);
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                CanHoDAL a = new CanHoDAL();
                a.MaCanHo = dr["MaCanHo"].ToString();
                a.Gia = (long)dr["Gia"];
                a.DienTich = (double)dr["DienTich"];
                a.TrangThai = (bool)dr["TrangThai"];
                a.SoPhong = (int)dr["SoPhong"];
                a.MaCuDan = dr["MaCuDan"].ToString();
                a.MaKhu = dr["MaKhu"].ToString();

                list.Add(a);
            }
            conn.Close();
            dgvCanHo.DataSource = list;
        }

        private void btnLamMoi_Click(object sender, EventArgs e)
        {
            txtMK.Enabled = true;
            txtMCH.Enabled = true;
            txtMCH.Clear();
            txtTimKiemCH.Clear();
            txtSP.Clear();
            txtDT.Clear();
            txtMCD.Clear();
            txtGia.Clear();
            txtMK.Clear();
            showCanHo();
        }

        //Them can ho
        private void button7_Click(object sender, EventArgs e)
        {
            SqlConnection conn = new SqlConnection(@"Data Source=DESKTOP-ELOR9UD\SQLEXPRESS;Initial Catalog=QuanLyChungCu2;Integrated Security=True");
            conn.Open();
            try
            {
                string query = "INSERT INTO CANHO (MaCanHo, DienTich, Gia,TrangThai,SoPhong,MaCuDan,MaKhu)VALUES ('" + txtMCH.Text + "','" + txtDT.Text + "','" + txtGia.Text + "','0','" + txtSP.Text + "',NULL,'" + txtMK.Text + "')";
                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.ExecuteNonQuery();
                MessageBox.Show("Thêm căn hộ thành công");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi:" + ex.Message);
            }
            conn.Close();
            showCanHo();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Login f = new Login();
            f.Show();
            this.Close();
        }

        private void btnSuaKCH_Click(object sender, EventArgs e)
        {
            SqlConnection conn = new SqlConnection(@"Data Source=DESKTOP-ELOR9UD\SQLEXPRESS;Initial Catalog=QuanLyChungCu2;Integrated Security=True");
            conn.Open();
            try
            {
                string query = "Update KHUCANHO set TenKhu = '" + txtTenKhu.Text + "',SoTang = '" + txtSoTang.Text + "',SoCanTT = '" + txtSoCan.Text + "',DiaChi = '" + txtDiaChi.Text + "' Where MaKhu='" + txtMaKhu.Text + "'";
                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.ExecuteNonQuery();
                MessageBox.Show("Sửa căn hộ thành công");
                showCanHo();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi:" + ex.Message);
            }
            conn.Close();
            
            
        }
        //Xoá khu căn hộ
        private void btnDelKCH_Click(object sender, EventArgs e)
        {
            SqlConnection conn = new SqlConnection(@"Data Source=DESKTOP-ELOR9UD\SQLEXPRESS;Initial Catalog=QuanLyChungCu2;Integrated Security=True");
            conn.Open();
            try
            {
                string query = "delete KHUCANHO where MaKhu='" + txtMK.Text + "'";
                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.ExecuteNonQuery();
                MessageBox.Show("Xoá căn hộ thành công");
                showCanHo();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi:" + ex.Message);
            }
            conn.Close();
        }

        //Sua can ho 
        private void button5_Click(object sender, EventArgs e)
        {
            SqlConnection conn = new SqlConnection(@"Data Source=DESKTOP-ELOR9UD\SQLEXPRESS;Initial Catalog=QuanLyChungCu2;Integrated Security=True");
            conn.Open();
            string query;
            if (txtMCD.Text.Length == 0)
            {
                query = "Update CANHO set DienTich = '" + txtDT.Text + "',Gia = '" + txtGia.Text + "',SoPhong = '" + txtSP.Text + "',MaKhu = '" + txtMK.Text + "',TrangThai='" + cbTrangThai.Checked + "' Where MaCanHo='" + txtMCH.Text + "'";
            }
            else
            {
                query = "Update CANHO set DienTich = '" + txtDT.Text + "',Gia = '" + txtGia.Text + "',SoPhong = '" + txtSP.Text + "',MaKhu = '" + txtMK.Text + "',MaCuDan='" + txtMCD.Text + "',TrangThai='" + cbTrangThai.Checked + "' Where MaCanHo='" + txtMCH.Text + "'";
            }

            try
            {

                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.ExecuteNonQuery();
                MessageBox.Show("Sửa căn hộ thành công");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi:" + ex.Message);
            }
            conn.Close();
            showCanHo();
        }

        //Xoa can ho
        private void button6_Click(object sender, EventArgs e)
        {
            SqlConnection conn = new SqlConnection(@"Data Source=DESKTOP-ELOR9UD\SQLEXPRESS;Initial Catalog=QuanLyChungCu2;Integrated Security=True");
            conn.Open();
            try
            {
                string query = "delete CANHO where MaCanHo='" + txtMCH.Text + "'";
                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.ExecuteNonQuery();
                MessageBox.Show("Xoá căn hộ thành công");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi:" + ex.Message);
            }
            conn.Close();
            showCanHo();
        }

        //Tim kiem can ho
        private void btnTimKiemCH_Click(object sender, EventArgs e)
        {
            List<CanHoDAL> list = new List<CanHoDAL>();
            SqlConnection conn = new SqlConnection(@"Data Source=DESKTOP-ELOR9UD\SQLEXPRESS;Initial Catalog=QuanLyChungCu2;Integrated Security=True");
            conn.Open();
            string query = "select * from CANHO where MaCanHo like '%" + txtTimKiemCH.Text + "%' or DienTich like'%" + txtTimKiemCH.Text + "%' or Gia like '%" + txtTimKiemCH.Text + "%' or SoPhong like'%" + txtTimKiemCH.Text + "%'or MaKhu like'%" + txtTimKiemCH.Text + "%'";
            SqlCommand cmd = new SqlCommand(query, conn);
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                CanHoDAL a = new CanHoDAL();
                a.MaCanHo = dr["MaCanHo"].ToString();
                a.Gia = (long)dr["Gia"];
                a.DienTich = (double)dr["DienTich"];
                a.TrangThai = (bool)dr["TrangThai"];
                a.SoPhong = (int)dr["SoPhong"];
                a.MaCuDan = dr["MaCuDan"].ToString();
                a.MaKhu = dr["MaKhu"].ToString();

                list.Add(a);
            }
            conn.Close();
            dgvCanHo.DataSource = list;
        }

        private void dgvThongTinMuaBan_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            int idx = e.RowIndex;
            txtMaHopDong.Enabled = false;
            txtMaHopDong.Text = dgvThongTinMuaBan.Rows[idx].Cells["MaHopDong"].Value.ToString();
            txtNgayGiaoDich.Text = dgvThongTinMuaBan.Rows[idx].Cells["NgayGiaoDich"].Value.ToString();
            txtDiaChiKhach.Text = dgvThongTinMuaBan.Rows[idx].Cells["DiaChiKH"].Value.ToString();
            txtTenKhachHang.Text = dgvThongTinMuaBan.Rows[idx].Cells["TenCuDan"].Value.ToString();
            txtCuDan.Text = dgvThongTinMuaBan.Rows[idx].Cells["MaCuDan"].Value.ToString();
            txtMaCanHo.Text = dgvThongTinMuaBan.Rows[idx].Cells["MaCanHo"].Value.ToString();
        }

        private void btnTimKiemHopDong_Click(object sender, EventArgs e)
        {
            List<LSMuaBanDAL> list = new List<LSMuaBanDAL>();
            SqlConnection conn = new SqlConnection(@"Data Source=DESKTOP-ELOR9UD\SQLEXPRESS;Initial Catalog=QuanLyChungCu2;Integrated Security=True");
            conn.Open();
            string query = "select * from HOPDONG where MaHopDong like '%" + txtTimKiemHopDong.Text + "%' or NgayGiaoDich like'%" + txtTimKiemHopDong.Text + "%' or DiaChiKH like '%" + txtTimKiemHopDong.Text + "%' or MaCuDan like'%" + txtTimKiemHopDong.Text + "%'or MaHopDong like'%" + txtTimKiemHopDong.Text + "%'";
            SqlCommand cmd = new SqlCommand(query, conn);
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                LSMuaBanDAL a = new LSMuaBanDAL();
                a.MaHopDong = dr["MaHopDong"].ToString();
                //a.TenKH = dr["TenCuDan"].ToString();
                a.MaCuDan = dr["MaCuDan"].ToString();
                a.DiaChiKH = dr["DiaChiKH"].ToString();
                a.NgayGiaoDich = dr["NgayGiaoDich"].ToString();
                list.Add(a);
            }
            conn.Close();
            dgvThongTinMuaBan.DataSource = list;
        }

        private void btnThemTKNhanVIen_Click(object sender, EventArgs e)
        {
            SqlConnection conn = new SqlConnection(@"Data Source=DESKTOP-ELOR9UD\SQLEXPRESS;Initial Catalog=QuanLyChungCu2;Integrated Security=True");
            conn.Open();
            try
            {
                string query = "INSERT INTO TAIKHOAN (TenTaiKhoan,MatKhau,VaiTro)VALUES ('" + txtTKNhanVien.Text + "','" + txtMKNhanVien.Text + "','0')";
                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.ExecuteNonQuery();
                MessageBox.Show("Thêm nhân viên thành công");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi:" + ex.Message);
            }
            conn.Close();
            showTK();
        }

        private void btnSuaTKNhanVIen_Click(object sender, EventArgs e)
        {
            SqlConnection conn = new SqlConnection(@"Data Source=DESKTOP-ELOR9UD\SQLEXPRESS;Initial Catalog=QuanLyChungCu2;Integrated Security=True");
            conn.Open();
            try
            {
                string query = "Update TaiKhoan set MatKhau = '" + txtMKNhanVien.Text + "' Where TenTaiKhoan='" + txtTKNhanVien.Text + "'";
                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.ExecuteNonQuery();
                MessageBox.Show("Sửa nhân viên thành công");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi:" + ex.Message);
            }
            conn.Close();
            showTK();
        }

        private void btnSuaTKAdmin_Click(object sender, EventArgs e)
        {
            SqlConnection conn = new SqlConnection(@"Data Source=DESKTOP-ELOR9UD\SQLEXPRESS;Initial Catalog=QuanLyChungCu2;Integrated Security=True");
            conn.Open();
            try
            {
                string query = "Update TaiKhoan set MatKhau = '" + txtTKAdmin.Text + "' Where TenTaiKhoan='" + txtMKAdmin.Text + "'";
                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.ExecuteNonQuery();
                MessageBox.Show("Sửa Admin thành công");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi:" + ex.Message);
            }
            conn.Close();
            showTK();
        }

        private void button13_Click(object sender, EventArgs e)
        {
            SqlConnection conn = new SqlConnection(@"Data Source=DESKTOP-ELOR9UD\SQLEXPRESS;Initial Catalog=QuanLyChungCu2;Integrated Security=True");
            conn.Open();
            try
            {
                string query = "delete TaiKhoan where TenTaiKhoan='" + txtTKNhanVien.Text + "'";
                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.ExecuteNonQuery();
                MessageBox.Show("Xoá thành công");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi:" + ex.Message);
            }
            conn.Close();
            showTK();
        }

        private void button16_Click(object sender, EventArgs e)
        {
            SqlConnection conn = new SqlConnection(@"Data Source=DESKTOP-ELOR9UD\SQLEXPRESS;Initial Catalog=QuanLyChungCu2;Integrated Security=True");
            conn.Open();
            try
            {
                string query = "delete TaiKhoan where TenTaiKhoan='" + txtTKAdmin.Text + "'";
                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.ExecuteNonQuery();
                MessageBox.Show("Xoá thành công");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi:" + ex.Message);
            }
            conn.Close();
            showTK();
        }

        private void button18_Click(object sender, EventArgs e)
        {
            SqlConnection conn = new SqlConnection(@"Data Source=DESKTOP-ELOR9UD\SQLEXPRESS;Initial Catalog=QuanLyChungCu2;Integrated Security=True");
            conn.Open();
            try
            {
                string query = "INSERT INTO TAIKHOAN (TenTaiKhoan,MatKhau,VaiTro)VALUES ('" + txtTKAdmin.Text + "','" + txtMKAdmin.Text + "','1')";
                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.ExecuteNonQuery();
                MessageBox.Show("Thêm thành công");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi:" + ex.Message);
            }
            conn.Close();
            showTK();
        }

        private void btnLamMoiTK_Click(object sender, EventArgs e)
        {
            
            txtTKAdmin.Clear();
            txtTKNhanVien.Clear();
            txtMKAdmin.Clear();
            txtMKNhanVien.Clear();

            showTK();
        }
        private void ToExcel(DataGridView dgvdgvThongTinMuaBan, string fileName)
        {
            //khai báo thư viện hỗ trợ Microsoft.Office.Interop.Excel
            Microsoft.Office.Interop.Excel.Application excel;
            Microsoft.Office.Interop.Excel.Workbook workbook;
            Microsoft.Office.Interop.Excel.Worksheet worksheet;
            try
            {
                //Tạo đối tượng COM.
                excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = false;
                excel.DisplayAlerts = false;
                //tạo mới một Workbooks bằng phương thức add()
                workbook = excel.Workbooks.Add(Type.Missing);
                worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets["Sheet1"];
                //đặt tên cho sheet
                worksheet.Name = "Quản lý Mua Bán";

                // export header trong DataGridView
                for (int i = 0; i < dgvThongTinMuaBan.ColumnCount; i++)
                {
                    worksheet.Cells[1, i + 1] = dgvThongTinMuaBan.Columns[i].HeaderText;
                }
                // export nội dung trong DataGridView
                for (int i = 0; i < dgvThongTinMuaBan.RowCount; i++)
                {
                    for (int j = 0; j < dgvThongTinMuaBan.ColumnCount; j++)
                    {
                        worksheet.Cells[i + 2, j + 1] = dgvThongTinMuaBan.Rows[i].Cells[j].Value.ToString();
                    }
                }
                // sử dụng phương thức SaveAs() để lưu workbook với filename
                workbook.SaveAs(fileName);
                //đóng workbook
                workbook.Close();
                excel.Quit();
                MessageBox.Show("Xuất dữ liệu ra Excel thành công!");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                workbook = null;
                worksheet = null;
            }
        }
        private void btnXuatFile_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                //gọi hàm ToExcel() với tham số là dtgDSHS và filename từ SaveFileDialog
                ToExcel(dgvThongTinCuDan, saveFileDialog1.FileName);
            }
        }
    }
}
