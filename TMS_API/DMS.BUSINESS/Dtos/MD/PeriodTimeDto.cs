using Common;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AutoMapper;
using DMS.CORE.Entities.MD;

namespace DMS.BUSINESS.Dtos.MD
{
    public class PeriodTimeDto : BaseMdDto, IDto, IMapFrom
    {
    
        [Key]
        public string Timeyear { get; set; }
        public bool? IsClosed { get; set; }
        public bool? IsDefault { get; set;}
        public string State { get => this.IsActive == true ? "Đang hoạt động" : "Khóa"; }
        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblMdPeriodTime, PeriodTimeDto>().ReverseMap();
        }
    }
}
