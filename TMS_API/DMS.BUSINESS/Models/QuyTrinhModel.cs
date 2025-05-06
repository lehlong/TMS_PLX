using DMS.BUSINESS.Dtos.MD;
using DMS.CORE.Entities.BU;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMS.BUSINESS.Models
{
    public class QuyTrinhModel
    {
        public TblBuCalculateDiscount header { get; set; } = new TblBuCalculateDiscount();
        public status Status { get; set; } = new status();
    }

    public class status
    {
        public string Code { set; get; }
        public string Content { set; get; }
        public string Link { set; get; }
    }
}
