using AutoMapper;
using Common;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DMS.CORE.Entities.MD;

namespace DMS.BUSINESS.Dtos.MD
{
   
        public class CurrencyDto : BaseMdDto, IMapFrom, IDto
        {

            [Description("STT")]
            public int OrdinalNumber { get; set; }

            [Key]
            [Description("Mã tiền tệ")]
            public string Code { get; set; }

            [Description("Tên đơn vị tiền tệ")]
            public string Name { get; set; }

            [Description("Tỷ giá")]
            public float Exchange_rate { get; set; }

            [Description("Trạng thái")]
            public string State { get => this.IsActive == true ? "Đang hoạt động" : "Khóa"; }

            public void Mapping(Profile profile)
            {
                profile.CreateMap<TblMdCurrency, CurrencyDto>().ReverseMap();
            }
        }
}

