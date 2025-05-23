﻿using AutoMapper;
using Common;
using DMS.BUSINESS.Dtos.AD;
using DMS.CORE.Entities.BU;
using System.ComponentModel.DataAnnotations;

namespace DMS.BUSINESS.Dtos.BU
{
    public class SmsDto : IMapFrom, IDto
    {
        [Key]
        public string Id { get; set; }

        public string? Status { get; set; }

        public string PhoneNumber { get; set; }

        public string Subject { get; set; }

        public string  Contents { get; set; }
        public string Is_send { get; set; }
        public int? NumberRerey { get; set; }
        public string HeaderId { get; set; }


        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblNotifySms, SmsDto>().ReverseMap();
        }
    }


}
