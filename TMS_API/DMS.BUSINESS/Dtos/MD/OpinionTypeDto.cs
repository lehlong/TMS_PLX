using AutoMapper;
using Common;
using DMS.CORE.Entities.MD;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMS.BUSINESS.Dtos.MD
{
    public class OpinionTypeDto : BaseMdDto, IDto, IMapFrom
    {
        [Description("STT")]
        public int OrdinalNumber { get; set; }
        [Key]
        [Description("Mã ý kiến")]
        public string Code { get; set; }
        [Description("Tên ý kiến")]
        public string Name { get; set; }

        [Description("Trạng thái")]
        public string State { get => this.IsActive == true ? "Đang hoạt động" : "Khóa"; }

        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblMdOpinionType, OpinionTypeDto>().ReverseMap();
        }

    }
}
