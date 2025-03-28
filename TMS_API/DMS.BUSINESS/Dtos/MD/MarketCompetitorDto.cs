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
    public class MarketCompetitorDto : BaseMdDto, IDto, IMapFrom
    {
        [Description("STT")]
        public int OrdinalNumber { get; set; }

        [Key]
        [Description("Mã thị trường")]
        public string Code { get; set; }

        [Description("Khoảng cách")]
        public decimal? Gap { get; set; }

        [Description("Mã thị trường")]
        public string? MarketCode { get; set; }

        [Description("Mã đối thủ")]
        public string? CompetitorCode { get; set; }

        [Description("Trạng thái")]
        public string State { get => this.IsActive == true ? "Đang hoạt động" : "Khóa"; }

        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblMdMarketCompetitor, MarketCompetitorDto>().ReverseMap();
        }

    }
}
