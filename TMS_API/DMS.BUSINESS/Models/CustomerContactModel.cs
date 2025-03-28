using DMS.BUSINESS.Dtos.MD;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMS.BUSINESS.Models
{
    public class CustomerContactModel
    {
        public string Customer_Code { get; set; }
        public List<CustomerContactDto> Contact_List { get; set; } = new List<CustomerContactDto>();
    }
}
