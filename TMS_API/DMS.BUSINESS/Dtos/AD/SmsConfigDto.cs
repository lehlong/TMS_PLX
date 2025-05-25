using AutoMapper;
using Common;
using DMS.CORE.Entities.AD;

using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace DMS.BUSINESS.Dtos.AD
{
    public class SmsConfigDto : IDto, IMapFrom
    {
        [Description("ID")]
        public int Id { get; set; }

        [Required]
        [Description("URL SMS")]
        public string UrlSms { get; set; }

        [Required]
        [Description("Tên đăng nhập")]
        public string Username { get; set; }

        [Required]
        [Description("Mật khẩu")]
        public string Password { get; set; }

        [Required]
        [Description("Mã CP")]
        public string CpCode { get; set; }

        [Required]
        [Description("Mã dịch vụ")]
        public string ServiceId { get; set; }

        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblSmsConfig, SmsConfigDto>().ReverseMap();
        }
    }
}
