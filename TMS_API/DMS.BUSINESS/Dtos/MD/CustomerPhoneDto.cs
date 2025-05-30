﻿using AutoMapper;
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
    public class CustomerPhoneDto : BaseMdDto, IDto, IMapFrom
    {
        [Description("STT")]
        public int OrdinalNumber { get; set; }

        [Key]
        [Description("Mã số điện thoại")]
        public string Code { get; set; }

        [Description("Mã khách hàng")]
        public string? CustomerCode { get; set; }

        [Description("Mã Thị trường")]
        public string? MarketCode { get; set; }

        [Description("Số điện thoại")]
        public string? Phone { get; set; }

        [Description("Trạng thái")]
        public string State { get => this.IsActive == true ? "Đang hoạt động" : "Khóa"; }

        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblMdCustomerPhone, CustomerPhoneDto>().ReverseMap();
        }
    }
}
