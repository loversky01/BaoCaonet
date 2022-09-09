using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BaoCaonet.Entity
{
    internal class ThongTinMuaBanDAL
    {
        string _maHopDong;
        string _tenCuDan;
        string _maCanHo;
        string _maCuDan;
        string _diaChiKH;
        long _gia;
        string _ngayGiaoDich;

        public string MaHopDong { get => _maHopDong; set => _maHopDong = value; }
        public string TenCuDan { get => _tenCuDan; set => _tenCuDan = value; }
        public string MaCanHo { get => _maCanHo; set => _maCanHo = value; }
        public string MaCuDan { get => _maCuDan; set => _maCuDan = value; }
        public string DiaChiKH { get => _diaChiKH; set => _diaChiKH = value; }
        public long Gia { get => _gia; set => _gia = value; }
        public string NgayGiaoDich { get => _ngayGiaoDich; set => _ngayGiaoDich = value; }
    }
}
