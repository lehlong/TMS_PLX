using AutoMapper;
using Common;
using DMS.CORE.Entities.MD;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMS.BUSINESS.Dtos.MD
{
    public class RetailPriceDto : BaseMdDto, IDto, IMapFrom
    {
        [Description("STT")]
        public int OrdinalNumber { get; set; }

        [Key]
        [Description("Mã")]
        public string Code { get; set; }

        [Description("Tên")]
        public string? GoodsCode { get; set; }

        [Description("Header code")]
        public string? GbllCode { get; set; }

        [Description("Giá cũ")]
        public decimal OldPrice { get; set; }

        [Description("Giá mới")]
        public decimal NewPrice { get; set; }

        [Description("Mức tăng so với vùng 1")]
        public decimal MucTangV1 { get; set; }


        [Description("Trạng thái")]
        public string State { get => this.IsActive == true ? "Đang hoạt động" : "Khóa"; }

        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblMdRetailPrice, RetailPriceDto>().ReverseMap();
        }
    }
}
