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
    public class MarketDto : BaseMdDto, IDto, IMapFrom
    {
        [Description("STT")]
        public int OrdinalNumber { get; set; }

        [Key]
        [Description("Mã thị trường")]
        public string Code { get; set; }

        [Description("Tên thị trường")]
        public string Name { get; set; }

        [Description("Khoảng cách")]
        public decimal? Gap { get; set; }

        [Description("hệ số")]
        public decimal? Coefficient { get; set; }
        
        [Description("Cước VC bình quân")]
        public decimal? CuocVCBQ { get; set; }

        [Description("CP chung chưa Cước VC")]
        public decimal? CPChungChuaCuocVC { get; set; }

        [Description("CK điều tiết cho xăng")]
        public decimal? CkDieuTietXang { get; set; }

        [Description("CK điều tiết cho dầu")]
        public decimal? CkDieuTietDau { get; set; }

        [Description("Mã vùng")]
        public string LocalCode { get; set; }

        [Description("Mã kho")]
        public string? WarehouseCode { get; set; }
        public string? FullName { get; set; }
        public string? Local2 { get; set; }

        [Description("Ngày tạo")]
        public DateTime? CreateDate { get; set; }

        [Description("Trạng thái")]
        public string State { get => this.IsActive == true ? "Đang hoạt động" : "Khóa"; }

        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblMdMarket, MarketDto>().ReverseMap();
        }

    }
}
