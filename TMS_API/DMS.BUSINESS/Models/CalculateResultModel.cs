using DMS.CORE.Entities.BU;
using DMS.CORE.Entities.MD;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMS.BUSINESS.Models
{
    public class CalculateResultModel
    {
        public List<PT>? PT { get; set; } = new List<PT>();
        public List<TblMdGoods>? lstGoods { get; set; }
        public DLG? DLG { get; set; } = new DLG();
        //public TblBuCalculateResultList HEADER_CR { get; set; } = new TblBuCalculateResultList();
        public List<DB>? DB { get; set; } = new List<DB>();
        public List<PT09>? PT09 { get; set; } = new List<PT09>();
        public List<PL1>? PL1 { get; set; } = new List<PL1>();
        public List<PL2>? PL2 { get; set; } = new List<PL2>();
        public List<PL3>? PL3 { get; set; } = new List<PL3>();
        public List<PL4>? PL4 { get; set; } = new List<PL4>();
        public List<FOB>? FOB { get; set; } = new List<FOB>();
        public List<VK11PT>? VK11PT { get; set; } = new List<VK11PT>();
        public List<VK11DB>? VK11DB { get; set; } = new List<VK11DB>();
        public List<VK11FOB>? VK11FOB { get; set; } = new List<VK11FOB>();
        public List<VK11TNPP>? VK11TNPP { get; set; } = new List<VK11TNPP>();
        public List<PTS> PTS { get; set; } = new List<PTS>();
        public List<VK11BB>? VK11BB { get; set; } = new List<VK11BB>();
        public List<BBDO>? BBDO { get; set; } = new List<BBDO>();
        public List<BBFO>? BBFO { get; set; } = new List<BBFO>();
        public List<Summary>? Summary { get; set; } = new List<Summary>();
    }
    public class DLG
    {
        public string? NameOld { get; set; }
        public List<DLG_1> Dlg_1 { get; set; } = new List<DLG_1>();
        public List<DLG_2> Dlg_2 { get; set; } = new List<DLG_2>();
        public List<DLG_3> Dlg_3 { get; set; } = new List<DLG_3>();
        public List<DLG_4> Dlg_4 { get; set; } = new List<DLG_4>();
        public List<DLG_4_Old> Dlg_4_Old { get; set; } = new List<DLG_4_Old>();
        public List<DLG_5> Dlg_5 { get; set; } = new List<DLG_5>();
        public List<DLG_6> Dlg_6 { get; set; } = new List<DLG_6>();
        public List<DLG_7> Dlg_7 { get; set; } = new List<DLG_7>();
        public List<DLG_8> Dlg_8 { get; set; } = new List<DLG_8>();
        public List<Dlg_TDGBL> Dlg_TDGBL { get; set; } = new List<Dlg_TDGBL>();
        public List<Dlg_TdGgptbl> Dlg_TdGgptbl { get; set; } = new List<Dlg_TdGgptbl>();

    }
    public class DLG_1
    {
        public string? Code { get; set; }
        public string? Col1 { get; set; }
        public decimal Col2 { get; set; } = 0;
        public decimal Col3 { get; set; } = 0;
        public decimal Col4 { get; set; } = 0;
        public decimal Col5 { get; set; } = 0;
        public decimal Col6 { get; set; } = 0;
        public decimal Col7 { get; set; } = 0;
    }
    public class DLG_2
    {
        public string? Code { get; set; }
        public string? Col1 { get; set; }
        public decimal Col2 { get; set; } = 0;
    }
    public class DLG_3
    {
        public string? Code { get; set; }
        public string? ColA { get; set; }
        public string? ColB { get; set; }
        public decimal Col1 { get; set; } = 0;
        public decimal Col2 { get; set; } = 0;
        public decimal Col3 { get; set; } = 0;
        public decimal Col4 { get; set; } = 0;
        public decimal Col5 { get; set; } = 0;
        public decimal Col6 { get; set; } = 0;
        public decimal Col7 { get; set; } = 0;
        public decimal Col8 { get; set; } = 0;
        public decimal Col9 { get; set; } = 0;
        public decimal Col10 { get; set; } = 0;
    }
    public class DLG_4
    {
        public string? Code { get; set; }
        public string? Type { get; set; }
        public string? ColA { get; set; }
        public string? ColB { get; set; }
        public decimal Col1 { get; set; } = 0;
        public decimal Col2 { get; set; } = 0;
        public decimal Col3 { get; set; } = 0;
        public decimal Col4 { get; set; } = 0;
        public decimal Col5 { get; set; } = 0;
        public decimal Col6 { get; set; } = 0;
        public decimal Col7 { get; set; } = 0;
        public decimal Col8 { get; set; } = 0;
        public decimal Col9 { get; set; } = 0;
        public decimal Col10 { get; set; } = 0;
        public decimal Col11 { get; set; } = 0;
        public decimal Col12 { get; set; } = 0;
        public decimal Col13 { get; set; } = 0;
        public decimal Col14 { get; set; } = 0;
        public decimal Col15 { get; set; } = 0;
        public decimal Col16 { get; set; } = 0;
        public bool IsBold { get; set; } = false;
    }
    public class DLG_4_Old
    {
        public string? Code { get; set; }
        public string? Type { get; set; }
        public string? ColA { get; set; }
        public string? ColB { get; set; }
        public decimal Col1 { get; set; } = 0;
        public decimal Col2 { get; set; } = 0;
        public decimal Col3 { get; set; } = 0;
        public decimal Col4 { get; set; } = 0;
        public decimal Col5 { get; set; } = 0;
        public decimal Col6 { get; set; } = 0;
        public decimal Col7 { get; set; } = 0;
        public decimal Col8 { get; set; } = 0;
        public decimal Col9 { get; set; } = 0;
        public decimal Col10 { get; set; } = 0;
        public decimal Col11 { get; set; } = 0;
        public decimal Col12 { get; set; } = 0;
        public decimal Col13 { get; set; } = 0;
        public decimal Col14 { get; set; } = 0;
        public decimal Col15 { get; set; } = 0;
        public decimal Col16 { get; set; } = 0;
        public bool IsBold { get; set; } = false;
    }
    public class DLG_5
    {
        public string? Code { get; set; }
        public string? ColA { get; set; }
        public string? ColB { get; set; }
        public decimal Col1 { get; set; } = 0;
        public decimal Col2 { get; set; } = 0;
        public decimal Col3 { get; set; } = 0;
        public decimal Col4 { get; set; } = 0;
        public decimal Col5 { get; set; } = 0;
    }
    public class DLG_6
    {
        public string? Code { get; set; }
        public string? ColA { get; set; }
        public string? ColB { get; set; }
        public decimal Col1 { get; set; } = 0;
        public decimal Col2 { get; set; } = 0;
        public decimal Col3 { get; set; } = 0;
        public decimal Col4 { get; set; } = 0;
        public decimal Col5 { get; set; } = 0;
        public decimal Col6 { get; set; } = 0;
        public decimal Col7 { get; set; } = 0;
        public string? Col8 { get; set; }

    }
    public class DLG_7
    {
        public string? Code { get; set; }
        public string? Type { get; set; }
        public string? ColA { get; set; }
        public decimal Col1 { get; set; } = 0;
        public decimal Col2 { get; set; } = 0;
        public decimal TangGiam1_2 { get; set; } = 0;
    }
    public class DLG_8
    {
        public string? Code { get; set; }
        public string? ColA { get; set; }
        public string? Type { get; set; }
        public decimal Col1 { get; set; } = 0;
        public decimal Col2 { get; set; } = 0;
        public decimal TangGiam1_2 { get; set; } = 0;
    }
    public class Dlg_TDGBL
    {
        public string? Code { get; set; }
        public string? ColA { get; set; }
        public decimal Col1 { get; set; } = 0;
        public decimal Col2 { get; set; } = 0;
        public decimal TangGiam1_2 { get; set; } = 0;
        public decimal Col3 { get; set; } = 0;
        public decimal Col4 { get; set; } = 0;
        public decimal TangGiam3_4 { get; set; } = 0;
    }

    public class Dlg_TdGgptbl
    {
        public string? Code { get; set; }
        public string? ColA { get; set; }
        public decimal Col1 { get; set; } = 0;
        public decimal Col2 { get; set; } = 0;
        public decimal TangGiam1_2 { get; set; } = 0;
    }

    public class PT
    {
        public string? Code { get; set; }
        public string? ColA { get; set; }
        public string? ColB { get; set; }
        public decimal Col1 { get; set; } = 0;
        public List<decimal> LG { get; set; } = new List<decimal>();
        public List<decimal> LN { get; set; } = new List<decimal>();
        public List<PT_GG> GG { get; set; } = new List<PT_GG>();
        public List<PT_BVMT> BVMT { get; set; } = new List<PT_BVMT>();
        public decimal Col3 { get; set; } = 0;
        public decimal Col4 { get; set; } = 0;
        public decimal Col5 { get; set; } = 0;
        public decimal Col6 { get; set; } = 0;
        public decimal Col7 { get; set; } = 0;
        public int Order { get; set; }
        public bool IsBold { get; set; }

    }
    public class PT_GG
    {
        public string? Code { get; set; }
        public decimal VAT { get; set; } = 0;
        public decimal NonVAT { get; set; } = 0;
    }
    public class PT_BVMT
    {
        public string? Code { get; set; }
        public decimal VAT { get; set; } = 0;
        public decimal NonVAT { get; set; } = 0;
    }
    public class DB
    {
        public string? Code { get; set; }
        public string? ColA { get; set; }
        public string? ColB { get; set; }
        public string? Col1 { get; set; }
        public decimal Col2 { get; set; } = 0;
        public List<decimal> LG { get; set; } = new List<decimal>();
        public decimal Col3 { get; set; } = 0;
        public decimal Col4 { get; set; } = 0;
        public decimal Col5 { get; set; } = 0;
        public decimal Col6 { get; set; } = 0;
        public decimal Col7 { get; set; } = 0;
        public decimal Col8 { get; set; } = 0;
        public decimal Col9 { get; set; } = 0;
        public decimal Col10 { get; set; } = 0;
        public List<DB_GG> GG { get; set; } = new List<DB_GG>();
        public List<decimal> LN { get; set; } = new List<decimal>();
        public bool IsBold { get; set; } = false;
        public List<DB_BVMT> BVMT { get; set; } = new List<DB_BVMT>();
    }

    public class DB_GG
    {
        public decimal VAT { get; set; } = 0;
        public decimal NonVAT { get; set; } = 0;
    }

    public class DB_BVMT
    {
        public string? Code { get; set; }
        public decimal VAT { get; set; } = 0;
        public decimal NonVAT { get; set; } = 0;
    }

    public class FOB
    {
        public string? Code { get; set; }
        public string? ColA { get; set; }
        public string? ColB { get; set; }
        public List<decimal> LG { get; set; } = new List<decimal>();
        public decimal Col1 { get; set; } = 0;
        public decimal Col2 { get; set; } = 0;
        public decimal Col3 { get; set; } = 0;
        public decimal Col4 { get; set; } = 0;
        public decimal Col5 { get; set; } = 0;
        public decimal Col6 { get; set; } = 0;
        public decimal Col7 { get; set; } = 0;
        public List<FOB_GG?> GG { get; set; } = new List<FOB_GG>();
        public decimal Col8 { get; set; } = 0;
        public List<decimal> LN { get; set; } = new List<decimal>();
        public bool IsBold { get; set; } = false;
        public List<PT_BVMT> BVMT { get; set; } = new List<PT_BVMT>();
    }

    public class FOB_GG
    {
        public decimal VAT { get; set; } = 0;
        public decimal NonVAT { get; set; } = 0;
    }

    public class PT09
    {
        public string? Code { get; set; }
        public string? ColA { get; set; }
        public string? ColB { get; set; }
        public List<decimal> LG { get; set; } = new List<decimal>();
        public decimal Col3 { get; set; } = 0;
        public decimal Col4 { get; set; } = 0;
        public decimal Col5 { get; set; } = 0;
        public decimal Col6 { get; set; } = 0;
        public decimal Col7 { get; set; } = 0;
        public decimal Col8 { get; set; } = 0;
        public List<PT09_GG> GG { get; set; } = new List<PT09_GG>();
        public decimal Col18 { get; set; }
        public List<decimal> LN { get; set; } = new List<decimal>();
        public List<PT_BVMT> BVMT { get; set; } = new List<PT_BVMT>();
        public bool IsBold { get; set; } = false;
    }
    public class PT09_GG
    {
        public decimal VAT { get; set; } = 0;
        public decimal NonVAT { get; set; } = 0;
    }
    public class BBDO
    {
        public string? Code { get; set; }
        public string? ColA { get; set; }
        public string? ColB { get; set; }
        public string? ColC { get; set; }
        public string? ColD { get; set; }
        public List<decimal> LG { get; set; } = new List<decimal>();
        public bool IsBold { get; set; } = false;
        public string? Col1 { get; set; }
        public string? Col2 { get; set; }
        public string? Col3 { get; set; }
        public string? Col4 { get; set; }
        public string? Col5 { get; set; }

        public decimal Col6 { get; set; } = 0;
        public decimal Col7 { get; set; } = 0;
        public decimal Col8 { get; set; } = 0;
        public decimal Col9 { get; set; }   = 0;
        public decimal Col10 { get; set; } = 0;
        public decimal Col11 { get; set; } = 0;
        public decimal Col12 { get; set; } = 0;
        public decimal Col13 { get; set; } = 0;
        public decimal Col14 { get; set; } = 0;
        public decimal Col15 { get; set; } = 0;
        public decimal Col16 { get; set; } = 0;
        public decimal Col17 { get; set; } = 0;
        public decimal Col18 { get; set; } = 0;
        public decimal Col19 { get; set; } = 0;
    }
    public class BBDO_MAP
    {
        public string? CustomerCode { get; set; }
        public string? CustomerName { get; set; }
        public string? PointCode { get; set; }
        public string? PointName { get; set; }
        public decimal CuocVcBq { get; set; } = 0;

    }
    public class BBFO
    {
        public string? Code { get; set; }
        public string? ColA { get; set; }
        public string? ColB { get; set; }
        public string? ColC { get; set; }
        public bool IsBold { get; set; } = false;
        public decimal Col1 { get; set; } = 0;
        public decimal Col2 { get; set; } = 0;
        public decimal Col3 { get; set; } = 0;
        public decimal Col4 { get; set; } = 0;
        public decimal Col5 { get; set; } = 0;
        public decimal Col6 { get; set; } = 0;
        public decimal Col7 { get; set; } = 0;
        public decimal Col8 { get; set; } = 0;
        public decimal Col9 { get; set; } = 0;
        public decimal Col10 { get; set; } = 0;
    }
    public class PL1
    {
        public string? Code { get; set; }
        public string? ColA { get; set; }
        public string? ColB { get; set; }
        public List<decimal> GG { get; set; } = new List<decimal>();
        public bool  IsBold { get; set; } = false;
    }
    public class PL2
    {
        public string? Code { get; set; }
        public string? ColA { get; set; }
        public string? ColB { get; set; }
        public List<decimal> GG { get; set; } = new List<decimal>();
        public bool IsBold { get; set; } = false;
    }

    public class PL3
    {
        public string? Code { get; set; }
        public string? ColA { get; set; }
        public string? ColB { get; set; }
        public List<decimal> GG { get; set; } = new List<decimal>();
        public bool IsBold { get; set; } = false;
    }
    public class PL4
    {
        public string? Code { get; set; }
        public string? ColA { get; set; }
        public string? ColB { get; set; }
        public List<decimal> GG { get; set; } = new List<decimal>();
        public bool IsBold { get; set; } = false;
    }
    public class VK11PT
    {
        public string? Code { get; set; }
        public string? ColA { get; set; }
        public string? ColB { get; set; }
        public bool IsBold { get; set; } = false;
        public string? Col1 { get; set; }
        public decimal Col2 { get; set; } = 0;
        public decimal Col3 { get; set; } = 0;
        public string? Col4 { get; set; }
        public string? Col5 { get; set; }
        public string? Col6 { get; set; }
        public string? Col7 { get; set; }
        public string? Col8 { get; set; }
        public decimal Col9 { get; set; } = 0;
        public string? Col10 { get; set; }
        public decimal Col11 { get; set; } = 0;
        public string? Col12 { get; set; }
        public string? Col13 { get; set; }
        public string? Col14 { get; set; }
        public string? Col15 { get; set; }
        public string? Col16 { get; set; }
        public string? Col17 { get; set; }
        public string? Col18 { get; set; }
    }
    public class VK11DB
    {
        public string? Code { get; set; }
        public string? ColA { get; set; }
        public string? ColB { get; set; }
        public string? ColC { get; set; }
        public bool IsBold { get; set; } = false;
        public string? Col1 { get; set; }
        public decimal Col2 { get; set; } = 0;
        public decimal Col3 { get; set; } = 0;
        public string? Col4 { get; set; }
        public string? Col5 { get; set; }
        public string? Col6 { get; set; }
        public string? Col7 { get; set; }
        public string? Col8 { get; set; }
        public decimal Col9 { get; set; } = 0;
        public string? Col10 { get; set; }
        public decimal Col11 { get; set; } = 0;
        public string? Col12 { get; set; }
        public string? Col13 { get; set; }
        public string? Col14 { get; set; }
        public string? Col15 { get; set; }
        public string? Col16 { get; set; }
        public string? Col17 { get; set; }
        public string? Col18 { get; set; }
    }
    public class VK11FOB
    {
        public string? Code { get; set; }
        public string? ColA { get; set; }
        public string? ColB { get; set; }
        public string? ColC { get; set; }
        public bool IsBold { get; set; } = false;
        public string? Col1 { get; set; }
        public decimal Col2 { get; set; } = 0;
        public decimal Col3 { get; set; } = 0;
        public string? Col4 { get; set; }
        public string? Col5 { get; set; }
        public string? Col6 { get; set; }
        public string? Col7 { get; set; }
        public string? Col8 { get; set; }
        public decimal Col9 { get; set; } = 0;
        public string? Col10 { get; set; }
        public decimal Col11 { get; set; } = 0;
        public string? Col12 { get; set; }
        public string? Col13 { get; set; }
        public string? Col14 { get; set; }
        public string? Col15 { get; set; }
        public string? Col16 { get; set; }
        public string? Col17 { get; set; }
        public string? Col18 { get; set; }
    }
    public class VK11TNPP
    {
        public string? Code { get; set; }
        public string? ColA { get; set; }
        public string? ColB { get; set; }
        public string? ColC { get; set; }
        public bool IsBold { get; set; } = false;
        public string? Col1 { get; set; }
        public decimal Col2 { get; set; } = 0;
        public decimal Col3 { get; set; } = 0;
        public string? Col4 { get; set; }
        public string? Col5 { get; set; }
        public string? Col6 { get; set; }
        public string? Col7 { get; set; }
        public string? Col8 { get; set; }
        public decimal Col9 { get; set; } = 0;
        public string? Col10 { get; set; }
        public decimal Col11 { get; set; } = 0;
        public string? Col12 { get; set; }
        public string? Col13 { get; set; }
        public string? Col14 { get; set; }
        public string? Col15 { get; set; }
        public string? Col16 { get; set; }
        public string? Col17 { get; set; }
        public string? Col18 { get; set; }
    }
    public class PTS
    {
        public string? ColA { get; set; }
        public string? Col1 { get; set; }
        public string? Col2 { get; set; }
        public string? Col3 { get; set; }
        public string? Col4 { get; set; }
        public string? Col5 { get; set; }
        public decimal Col6 { get; set; } = 0;
        public decimal Col7 { get; set; } = 0;
        public string? Col8 { get; set; }
        public string? Col9 { get; set; }
        public string? Col10 { get; set; }
        public string? Col11 { get; set; }
        public bool IsBold { get; set; } = false;
    }
    public class VK11BB
    {
        public string? Code { get; set; }
        public string? ColA { get; set; }
        public string? ColB { get; set; }
        public string? ColC { get; set; }
        public string? ColD { get; set; }
        public List<decimal> LG { get; set; } = new List<decimal>();
        public bool IsBold { get; set; } = false;
        public string? Col1 { get; set; }
        public string? Col2 { get; set; }
        public string? Col3 { get; set; }
        public string? Col4 { get; set; }
        public string? Col5 { get; set; }

        public decimal Col6 { get; set; } = 0;
        public string? Col7 { get; set; }
        public string? Col8 { get; set; }
        public string? Col9 { get; set; }
        public string? Col10 { get; set; }
        public string? Col11 { get; set; }
        public string? Col12 { get; set; }
        public string? Col13 { get; set; }
        public string? Col14 { get; set; }
        public string? Col15 { get; set; }
    }
    public class Summary
    {
        public string? Code { get; set; }
        public string? ColA { get; set; }
        public string? ColB { get; set; }
        public string? ColC { get; set; }
        public string? ColD { get; set; }
        public List<decimal> LG { get; set; } = new List<decimal>();
        public bool IsBold { get; set; } = false;
        public string? Col1 { get; set; }
        public string? Col2 { get; set; }
        public string? Col3 { get; set; }
        public string? Col4 { get; set; }
        public string? Col5 { get; set; }

        public decimal Col6 { get; set; } = 0;
        public string? Col7 { get; set; }
        public string? Col8 { get; set; }
        public string? Col9 { get; set; }
        public string? Col10 { get; set; }
        public string? Col11 { get; set; }
        public string? Col12 { get; set; }
        public string? Col13 { get; set; }
        public string? Col14 { get; set; }
        public string? Col15 { get; set; }
    }
    public class QueryModel
    {
        public string? FDate { get; set; }
    }


}
