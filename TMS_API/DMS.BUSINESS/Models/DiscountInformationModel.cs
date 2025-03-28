using DMS.CORE.Entities.BU;
using DMS.CORE.Entities.MD;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMS.BUSINESS.Models
{
    public class DiscountInformationModel
    {
        public List<TblMdGoods> lstGoods { get; set; }
        public List<TblMdCompetitor> lstCompetitor { get; set; }
        public List<discout> discount { get; set; } = new List<discout>();
        public List<TblBuDiscountInformationList> lstDIL { get; set; } = new List<TblBuDiscountInformationList>();
    }

    public class discout
    {
        public string colA { get; set; }
        public string colB { get; set; }
        public decimal? col1 { get; set; }
        public List<decimal?> gaps { get; set; } = new List<decimal?>();
        public decimal? col4 { get; set; }
        public List<decimal?> cuocVCs { get; set; } = new List<decimal?>();
        public bool IsBold { get; set; } = false;
        public List<CK> CK { get; set; } = new List<CK>();

    }
    public class CK
    {
        public decimal? plxna { get; set; }
        public List<DT> DT { get; set; } = new List<DT>();
    }

    public class DT
    {
        public List<decimal?> ckCl { get; set; } = new List<decimal?>();
    }

    //public class CkCl
    //{
    //    public decimal? ck { get; set; }
    //    public decimal? cl { get; set; }
    //}
}
