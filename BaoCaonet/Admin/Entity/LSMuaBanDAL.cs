using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BaoCaonet.Entity
{
    internal class LSMuaBanDAL
    {
        string _tenKH;
        string _maHopDong;
        string _ngayGiaoDich;
        string _maCuDan;
        string _diaChiKH;


        public string MaHopDong { get => _maHopDong; set => _maHopDong = value; }
        public string TenKH { get => _tenKH; set => _tenKH = value; }
        public string MaCuDan { get => _maCuDan; set => _maCuDan = value; }
        public string DiaChiKH { get => _diaChiKH; set => _diaChiKH = value; }
        public string NgayGiaoDich { get => _ngayGiaoDich; set => _ngayGiaoDich = value; }
    }
}
