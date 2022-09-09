using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BaoCaonet.Entity
{
     class ThongTinCuDanDAL
    {
        string _maCuDan;
        string _tenCuDan;
        string _ngaySinh;
        string _gioiTinh;
        string _soDT;
        string _soCMT;
        string _queQuan;

        public string MaCuDan { get => _maCuDan; set => _maCuDan = value; }
        public string TenCuDan { get => _tenCuDan; set => _tenCuDan = value; }
        public string NgaySinh { get => _ngaySinh; set => _ngaySinh = value; }
        public string GioiTinh { get => _gioiTinh; set => _gioiTinh = value; }
        public string SoDT { get => _soDT; set => _soDT = value; }
        public string SoCMT { get => _soCMT; set => _soCMT = value; }
        public string QueQuan { get => _queQuan; set => _queQuan = value; }
    }
}
