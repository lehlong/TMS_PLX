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
    public class CustomerDto : BaseMdDto, IDto, IMapFrom
    {
        [Description("STT")]
        public int OrdinalNumber { get; set; }

        [Key]
        [Description("Mã")]
        public string Code { get; set; }

        [Description("Tên")]
        public string Name { get; set; }

        //[Description("Số điện thoại")]
        //public string Phone { get; set; }

        //[Description("email")]
        //public string Email { get; set; }

        [Description("địa chỉ")]
        public string Address { get; set; }

        [Description("Hạn thanh toán")]
        public string PaymentTerm { get; set; }

        [Description("Cự ly BQ")]
        public decimal Gap { get; set; }

        [Description("Cước vận chuyển bình quân")]
        public decimal CuocVcBq { get; set; }

        [Description("Hỗ trợ cước Vân chuyển")]
        public decimal HoTroCuocVc { get; set; }

        [Description("Mức giảm xăng")]
        public decimal MgglhXang { get; set; }

        [Description("Mức giảm dầu")]
        public decimal MgglhDau { get; set; }

        [Description("thông tin mua hàng")]
        public string BuyInfo { get; set; }

        [Description("Lãi vay ngân hàng")]
        public decimal BankLoanInterest { get; set; }

        [Description("Mã phương thức mua hàng")]
        public string SalesMethodCode { get; set; }

        [Description("Mã Vùng")]
        public string LocalCode { get; set; }

        [Description("Mã kiểu khách hàng")]
        public string CustomerTypeCode { get; set; }

        [Description("Mã thị trường")]
        public string MarketCode { get; set; }

        [Description("Đề xuất mức giảm giá tại kho bên bán (FOB)")]
        public decimal Fob { get; set; }

        [Description("Trạng thái")]
        public string State { get => this.IsActive == true  ? "Đang hoạt động" : "Khóa"; }

        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblMdCustomer, CustomerDto>().ReverseMap();
        }
    }
}
