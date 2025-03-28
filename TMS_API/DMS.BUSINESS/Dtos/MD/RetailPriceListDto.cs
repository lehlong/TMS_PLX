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
    class RetailPriceListDto : BaseMdDto, IDto, IMapFrom
    {
        [Description("STT")]
        public int OrdinalNumber { get; set; }

        [Key]
        [Description("Mã")]
        public string Code { get; set; }

        [Description("Tên")]
        public string? Name { get; set; }

        [Description("Ngày bắt tạo")]
        public DateTime? FDate { get; set; }

        [Description("Ngày hết hạn")]
        public DateTime? TDate { get; set; }

        [Description("Trạng thái")]
        public string State { get => this.IsActive == true ? "Đang hoạt động" : "Khóa"; }

        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblMdRetailPriceList, RetailPriceListDto>().ReverseMap();
        }
    }
}
