using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BaoCaonet.Entity
{
    internal class TaiKhoanDAL
    {
        string _tenTaiKhoan;
        string _matKhau;
        string _vaiTro;

        public string TenTaiKhoan { get => _tenTaiKhoan; set => _tenTaiKhoan = value; }
        public string MatKhau { get => _matKhau; set => _matKhau = value; }
        public string VaiTro { get => _vaiTro; set => _vaiTro = value; }
    }
}
