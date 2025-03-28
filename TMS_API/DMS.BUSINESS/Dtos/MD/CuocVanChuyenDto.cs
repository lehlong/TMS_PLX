using AutoMapper;
using DMS.CORE.Entities.MD;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Common;
using System.ComponentModel.DataAnnotations.Schema;

namespace DMS.BUSINESS.Dtos.MD
{
    public class CuocVanChuyenDto : BaseMdDto, IMapFrom, IDto
    {

        [Description("STT")]
        public int OrdinalNumber { get; set; }

        [Key]
        [Description("Mã cước vận chuyển")]
        public string? Code { set; get; }

        [Description("List code")]
        public string? HeaderCode { set; get; }

        [Description("Loại hình vận chuyển")]
        public string? VSART { get; set; }

        [Description("Mã đơn vị")]
        public string? TDLNR { get; set; }

        [Description("Mã kho")]
        public string? KNOTA { get; set; }

        [Description("Nơi nhận hàng")]
        public string? OIGKNOTD { get; set; }

        [Description("Mặt hàng")]
        public string? MATNR { get; set; }

        [Description("Đơn vị bán hàng")]
        public string? VRKME { get; set; }

        [Description("Giá")]
        public decimal? KBETR { get; set; }

        [Description("Đơn vị tiền tệ")]
        public string? KONWA { get; set; }

        [Description("")]
        public string? KPEIN { get; set; }

        [Description("")]
        public string? KMEIN { get; set; }

        [Description("Ngày bắt đầu")]
        public string? DATAB { get; set; }

        [Description("Ngày kết thúc")]
        public string? DATBI { get; set; }

        [Description("STT")]
        public int? RowIndex { get; set; }

        [Description("Trạng thái")]
        public string? State { get => this.IsActive == true ? "Đang hoạt động" : "Khóa"; }

        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblMdCuocVanChuyen, CuocVanChuyenDto>().ReverseMap();
        }

    }
}
