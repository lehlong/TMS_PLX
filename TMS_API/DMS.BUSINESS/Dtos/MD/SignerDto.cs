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
    public class SignerDto: BaseMdDto, IMapFrom, IDto
    {
        [Description("STT")]
        public int OrdinalNumber { get; set; }

        [Key]
        [Description("Mã")]
        public string Code { set; get; }

        [Description("Chức vụ")]
        public string Position { get; set; }

        [Description("Tên người ký")]
        public string Name { get; set; }

        [Description("Mặc Định")]
        public bool? IsSelect { get; set; }

        [Description("kiểu")]
        public string Type { get; set; }

        [Description("Trạng thái")]
        public string State { get => this.IsActive == true ? "Đang hoạt động" : "Khóa"; }

        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblMdSigner, SignerDto>().ReverseMap();
        }
    }
}
