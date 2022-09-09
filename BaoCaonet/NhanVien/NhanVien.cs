using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;
using System.IO;
namespace BaoCaonet.NhanVien
{
    public partial class NhanVien : Form
    {
        public NhanVien()
        {
            InitializeComponent();
            Lichsumuaban();
            QuanLyCuDan();
            ThongTinMuaBan();
            ThongTinCanHo();
            QuanLyHopDong();
        }

        //=============== Lịch Sử Mua Bán ===================
        public void Lichsumuaban()
        {
            string sql = "select * from lichsu";
            DataTable tb = KetNoi.SelectDB(sql);
            dgvLSMuaBan.DataSource = tb;
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewRow myrow = new DataGridViewRow();
            myrow = dgvLSMuaBan.Rows[e.RowIndex];
            txtTenKH.Text = myrow.Cells["tenkh"].Value.ToString();
            txtMHD.Text = myrow.Cells["mahopdong"].Value.ToString();
            dtNgayGD.Value = DateTime.Parse(myrow.Cells["ngaygiaodich"].Value.ToString());
            txtDCKH.Text = myrow.Cells["diachikh"].Value.ToString();
            txtMCD.Text = myrow.Cells["macudan"].Value.ToString();
        }




        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void label9_Click(object sender, EventArgs e)
        {

        }
        //=================== Kết thúc Lịch Sử Mua Bán ========================

        //=================== Quản Lý Cư Dân ==================================
        public void QuanLyCuDan()
        {
            string sql = "select * from cudan";
            DataTable tb = KetNoi.SelectDB(sql);
            dgvQLCD.DataSource = tb;
        }

        private void dgvQLCD_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewRow myrow = new DataGridViewRow();
            myrow = dgvQLCD.Rows[e.RowIndex];
            txt_QLCD_MCD.Text = myrow.Cells["QLCD_macudan"].Value.ToString();
            txtTCD.Text = myrow.Cells["tencudan"].Value.ToString();
            dtNgaysinh.Value = DateTime.Parse(myrow.Cells["ngaysinh"].Value.ToString());
            txtGioiTinh.Text = myrow.Cells["gioitinh"].Value.ToString();
            txtSDT.Text = myrow.Cells["sodt"].Value.ToString();
            txtSoCMT.Text = myrow.Cells["socmt"].Value.ToString();
            txtQue.Text = myrow.Cells["quequan"].Value.ToString();
        }

        private void btn_QLCD_Them_Click(object sender, EventArgs e)
        {
            if (txt_QLCD_MCD.Text == "")
                MessageBox.Show("Chưa Nhập Mã Cư Dân cần thêm!");
            else
            {
                string sql = "INSERT INTO CUDAN (MaCuDan, TenCuDan, NgaySinh,GioiTinh,SoDT,SoCMT,QueQuan)VALUES ('" + txt_QLCD_MCD.Text + "',N'" + txtTCD.Text + "','" + dtNgaysinh.Value.ToString("yyyy-MM-dd") + "','" + txtGioiTinh.Text + "','" + txtSDT.Text + "','" + txtSoCMT.Text + "',N'" + txtQue.Text + "')";
                KetNoi.UpInsDelDB(sql);
                MessageBox.Show("Thêm cư dân thành công!");
                QuanLyCuDan();
            }
        }

        private void btn_QLCD_Sua_Click(object sender, EventArgs e)
        {
            string sql = "update cudan set tencudan=N'" + txtTCD.Text + "',ngaysinh='" + dtNgaysinh.Value.ToString("yyyy-MM-dd") + "',gioitinh='" + txtGioiTinh.Text + "',sodt='" + txtSDT.Text + "',socmt='" + txtSoCMT.Text + "',quequan=N'" + txtQue.Text + "' where macudan='" + txt_QLCD_MCD.Text + "'";
            KetNoi.UpInsDelDB(sql);
            MessageBox.Show("Cập nhật thành công!");
            QuanLyCuDan();
        }

        private void btn_QLCD_Xoa_Click(object sender, EventArgs e)
        {
            string sql = "delete from cudan where macudan = '" + txt_QLCD_MCD.Text + "'";
            KetNoi.UpInsDelDB(sql);
            MessageBox.Show("Xóa thành công!");
            QuanLyCuDan();
        }

        private void btn_QLCD_TimKiem_Click(object sender, EventArgs e)
        {
            string sql = "select * from cudan where macudan = '" + txt_QLCD_TK.Text + "' or tencudan like N'%" + txt_QLCD_TK.Text + "%' or sodt = '" + txt_QLCD_TK.Text + "' or socmt = '" + txt_QLCD_TK.Text + "' or quequan like N'%" + txt_QLCD_TK.Text + "%'";
            DataTable tb = KetNoi.SelectDB(sql);
            dgvQLCD.DataSource = tb;
        }

        private void btn_QLCD_Reset_Click(object sender, EventArgs e)
        {
            txt_QLCD_TK.Text = "";
            txt_QLCD_MCD.Text = "";
            txtTCD.Text = "";
            dtNgaysinh.Value = DateTime.Now;
            txtGioiTinh.Text = "";
            txtSDT.Text = "";
            txtSoCMT.Text = "";
            txtQue.Text = "";
        }

        private void btnTimkiem_Click_1(object sender, EventArgs e)
        {
            string sql = "select * from lichsu where tenkh like N'%" + txtTK.Text + "%' or diachikh like N'%" + txtTK.Text + "%' or mahopdong = '" + txtTK.Text + "' or macudan = '" + txtTK.Text + "'";
            DataTable tb = KetNoi.SelectDB(sql);
            dgvLSMuaBan.DataSource = tb;
        }
        //================== Thông tin Mua Bán ===============================
        public void ThongTinMuaBan()
        {
            string sql = "select hopdong.*,cudan.tencudan,gia from hopdong,cudan,canho where hopdong.macudan=cudan.macudan and hopdong.macanho=canho.macanho";
            DataTable tb = KetNoi.SelectDB(sql);
            dgvTTMB.DataSource = tb;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            string sql = "select hopdong.*,cudan.tencudan,gia from hopdong,cudan,canho where hopdong.macudan=cudan.macudan and hopdong.macanho=canho.macanho and(hopdong.mahopdong='" + txt_TTMB_TK.Text + "' or diachikh like N'%" + txt_TTMB_TK.Text + "%' or cudan.tencudan like N'%" + txt_TTMB_TK.Text + "%' or hopdong.macanho='" + txt_TTMB_TK.Text + "' or hopdong.macudan='" + txt_TTMB_TK.Text + "')";
            DataTable tb = KetNoi.SelectDB(sql);
            dgvTTMB.DataSource = tb;
        }
        //================== Kết Thúc Thông Tin Mua Bán ======================= 

        //================== Thông Tin Hợp Đồng ===============================
        public void ThongTinCanHo()
        {
            string sql = "select macanho,dientich,gia,sophong,tenkhu from canho,khucanho where canho.makhu=khucanho.makhu and trangthai = 'false' ";
            DataTable tb = KetNoi.SelectDB(sql);
            dgvTTCH.DataSource = tb;
        }
        private void btn_TTHD_Click(object sender, EventArgs e)
        {
            //string sql1 = "select * from cudan";
            //DataTable tb = KetNoi.SelectDB(sql1);
            //string s = tb.Rows[tb.Rows.Count - 1]["macudan"].ToString();//lấy dòng cuối xử lí
            //int macudan = Convert.ToInt32(s);
            //macudan++;
            //string inscd = "insert into cudan values('" + macudan + "',N'" + txt_TTCH_tenkh.Text + "','" + dt_TTCH_ngaysinh.Value.ToString("yyyy-MM-dd") + "','" + txt_TTCH_gioitinh.Text + "','" + txt_TTCH_sdt.Text + "','" + txt_TTCH_cmt.Text + "',N'" + txt_TTCH_dc.Text + "')";
            //KetNoi.UpInsDelDB(inscd);
            //QuanLyCuDan();

            //string upch = "update canho set trangthai='true' where macanho='" + txt_TTCH_mach.Text + "'";
            //KetNoi.UpInsDelDB(upch);
            //ThongTinCanHo();

            //string sql2 = "select * from hopdong";
            //DataTable tb2 = KetNoi.SelectDB(sql2);
            //string s2 = tb2.Rows[tb2.Rows.Count - 1]["mahopdong"].ToString();//lấy dòng cuối xử lí
            //string s1 = "";
            //for (int i = 2; i < s2.Length; i++)
            //{
            //    s1 = s1 + s2[i];
            //}
            ////
            //int a2 = 0;
            //string ms = "";
            //a2 = Convert.ToInt32(s1);
            //if (a2 >= 9)
            //{
            //    ms = "HD00000000" + (a2 + 1).ToString();
            //}
            //else
            //{
            //    ms = "HD000000000" + (a2 + 1).ToString();
            //}
            //string inshd = "insert into hopdong values('" + ms + "','" + dt_TTCH_ngayGD.Value.ToString("yyyy-MM-dd") + "',N'" + txt_TTCH_dc.Text + "','" + macudan + "','" + txt_TTCH_mach.Text + "')";
            //KetNoi.UpInsDelDB(inshd);
            //QuanLyHopDong();

            //string insls = "insert into lichsu values(N'" + txt_TTCH_tenkh.Text + "','" + ms + "','" + dt_TTCH_ngayGD.Value.ToString("yyyy-MM-dd") + "',N'" + txt_TTCH_dc.Text + "','" + txt_TTCH_mach.Text + "')";
            //KetNoi.UpInsDelDB(insls);
            //Lichsumuaban();
            //ThongTinMuaBan();
            //MessageBox.Show("Đăng kí thành công!");
            CreateWordDocument(@"D:\.net(C#)\BaoCaonet\BaoCaonet\word\temp.docx", @"D:\.net(C#)\BaoCaonet\BaoCaonet\word\hopdong.docx");
            MessageBox.Show("Xuất file thành công");
        }

        private void dgvTTCH_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            txt_TTCH_mach.Enabled = false;
            txt_TTCH_gia.Enabled = false;
            DataGridViewRow myrow = new DataGridViewRow();
            myrow = dgvTTCH.Rows[e.RowIndex];
            txt_TTCH_mach.Text = myrow.Cells["TTCH_macanho"].Value.ToString();
            txt_TTCH_gia.Text = myrow.Cells["TTCH_gia"].Value.ToString();
        }

        private void dgvQLHD_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        //====================== Kết Thúc Thông Tin Hợp Đồng =================

        //====================== Quản Lý Căn Hộ ====================
        public void QuanLyHopDong()
        {
            string sql = "select hopdong.*,cudan.tencudan,gia from hopdong,cudan,canho where hopdong.macudan=cudan.macudan and hopdong.macanho=canho.macanho";
            DataTable tb = KetNoi.SelectDB(sql);
            dgvQLHD.DataSource = tb;
        }
        private void btnSua_Click(object sender, EventArgs e)
        {
            string sql = "update cudan set tencudan=N'" + txtTenKH.Text + "' where macudan='" + txtMCD.Text + "'";
            KetNoi.UpInsDelDB(sql);

            string sql2 = "insert into lichsu values(N'" + txtTenKH.Text + "','" + txtMHD.Text + "','" + dtNgayGD.Value.ToString("yyyy-MM-dd") + "',N'" + txtDCKH.Text + "','" + txtMCD.Text + "')";
            KetNoi.UpInsDelDB(sql2);
            MessageBox.Show("Cập nhật thành công!");
            QuanLyHopDong();
            QuanLyCuDan();
            Lichsumuaban();
        }
        private void btnReset_Click(object sender, EventArgs e)
        {
            txtTenKH.Text = "";
            txtDCKH.Text = "";
            txtMCD.Text = "";
            txtTK.Text = "";
            txtMHD.Text = "";
            dtNgayGD.Value = DateTime.Now;
        }

        private void dgvQLHD_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            txtMCD.Enabled = false;
            DataGridViewRow myrow = new DataGridViewRow();
            myrow = dgvQLHD.Rows[e.RowIndex];
            txtMHD.Text = myrow.Cells["QLHD_mahopdong"].Value.ToString();
            txtTenKH.Text = myrow.Cells["QLHD_tenkh"].Value.ToString();
            txtDCKH.Text = myrow.Cells["QLHD_dckh"].Value.ToString();
            txtMCD.Text = myrow.Cells["QLHD_macudan"].Value.ToString();
            dtNgayGD.Value = DateTime.Parse(myrow.Cells["QLHD_ngayGD"].Value.ToString());
        }

        private void FindAndReplace(Word.Application wordApp, object ToFindText, object replaceWithText)
        {
            object matchCase = true;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundLike = false;
            object nmatchAllforms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiactitics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;

            wordApp.Selection.Find.Execute(ref ToFindText,
                ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundLike,
                ref nmatchAllforms, ref forward,
                ref wrap, ref format, ref replaceWithText,
                ref replace, ref matchKashida,
                ref matchDiactitics, ref matchAlefHamza,
                ref matchControl);
        }

        //Creeate the Doc Method
        private void CreateWordDocument(object filename, object SaveAs)
        {
            Word.Application wordApp = new Word.Application();
            object missing = Missing.Value;
            Word.Document myWordDoc = null;

            if (File.Exists((string)filename))
            {
                object readOnly = false;
                object isVisible = false;
                wordApp.Visible = false;

                myWordDoc = wordApp.Documents.Open(ref filename, ref missing, ref readOnly,
                                        ref missing, ref missing, ref missing,
                                        ref missing, ref missing, ref missing,
                                        ref missing, ref missing, ref missing,
                                        ref missing, ref missing, ref missing, ref missing);
                myWordDoc.Activate();

                //find and replace
                this.FindAndReplace(wordApp, "<TenCuDan>", txt_TTCH_tenkh.Text);
                this.FindAndReplace(wordApp, "<NgaySinh>", dt_TTCH_ngaysinh.Value.ToShortDateString());
                this.FindAndReplace(wordApp, "<SoCMT>", txt_TTCH_cmt.Text);
                this.FindAndReplace(wordApp, "<QueQuan>", txt_TTCH_dc.Text);
                this.FindAndReplace(wordApp, "<SoDT>", txt_TTCH_dc.Text);
            }
            else
            {
                MessageBox.Show("File not Found!");
            }

            //Save as
            myWordDoc.SaveAs2(ref SaveAs, ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing);

            myWordDoc.Close();
            wordApp.Quit();
            MessageBox.Show("File Created!");

        }
    }
}
