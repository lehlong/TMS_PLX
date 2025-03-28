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
    public class UnitDto : BaseMdDto, IDto, IMapFrom
    {
        [Description("STT")]
        public int OrdinalNumber { get; set; }

        [Key]
        [Description("Mã đơn vị tính")]
        public string Code { get; set; }

        [Description("Tên đơn vị tính")]
        public string Name { get; set; }


        [Description("Trạng thái")]
        public string State { get => this.IsActive == true ? "Đang hoạt động" : "Khóa"; }

        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblMdUnit, UnitDto>().ReverseMap();
        }
    
    }
}
