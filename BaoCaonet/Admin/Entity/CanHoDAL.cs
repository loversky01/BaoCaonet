using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BaoCaonet.Entity
{
     class CanHoDAL
    {
        string _maCanHo;
        double _dienTich;
        long _gia;
        bool _trangThai;
        int _soPhong;
        string _maCuDan;
        string _maKhu;

        public string MaCanHo { get => _maCanHo; set => _maCanHo = value; }
        public double DienTich { get => _dienTich; set => _dienTich = value; }
        public long Gia { get => _gia; set => _gia = value; }
        public bool TrangThai { get => _trangThai; set => _trangThai = value; }
        public int SoPhong { get => _soPhong; set => _soPhong = value; }
        public string MaCuDan { get => _maCuDan; set => _maCuDan = value; }
        public string MaKhu { get => _maKhu; set => _maKhu = value; }
    }
}
