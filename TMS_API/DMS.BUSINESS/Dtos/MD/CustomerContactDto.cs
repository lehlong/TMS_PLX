using AutoMapper;
using DMS.CORE.Entities.MD;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Common;

namespace DMS.BUSINESS.Dtos.MD
{
    public class CustomerContactDto : BaseMdDto, IMapFrom, IDto
    {
        [Description("STT")]
        public int OrdinalNumber { get; set; }

        [Key]
        [Description("Mã tiền tệ")]
        public string Code { get; set; }

        [Description("Mã khách hàng")]
        public string Customer_Code { get; set; }

        [Description("Kiểu liên hệ")]
        public string Type { get; set; }

        [Description("Giá trị")]
        public string Value { get; set; }

        [Description("Trạng thái")]
        public string State { get => this.IsActive == true ? "Đang hoạt động" : "Khóa"; }

        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblMdCustomerContact, CustomerContactDto>().ReverseMap();
        }
    }
}
