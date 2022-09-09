using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BaoCaonet.Entity
{
     class KhuCanHoDAL
    {
        string _maKhu;
        string _tenKhu;
        int _soTang;
        int _soCanTT;
        string _diaChi;

        public string MaKhu { get => _maKhu; set => _maKhu = value; }
        public string TenKhu { get => _tenKhu; set => _tenKhu = value; }
        public int SoTang { get => _soTang; set => _soTang = value; }
        public int SoCanTT { get => _soCanTT; set => _soCanTT = value; }
        public string DiaChi { get => _diaChi; set => _diaChi = value; }
    }
}
