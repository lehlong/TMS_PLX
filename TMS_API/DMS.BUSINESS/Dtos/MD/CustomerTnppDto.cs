using AutoMapper;
using Common;
using DMS.CORE.Entities.MD;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMS.BUSINESS.Dtos.MD
{
    public class CustomerTnppDto : BaseMdDto, IDto, IMapFrom
    {
        [Description("STT")]
        public int OrdinalNumber { get; set; }

        [Key]
        [Description("Mã")]
        public string Code { get; set; }

        [Description("Tên")]
        public string Name { get; set; }

        [Description("mã vùng")]
        public string LocalCode { get; set; }

        [Description("mã thị trường")]
        public string MarketCode { get; set; }

        [Description("Cự ly bình quân")]
        public decimal CuLyBq { get; set; }

        [Description("Kinh phí chung chưa Cước vân chuyển")]
        public decimal Cpccvc { get; set; }

        [Description("Cước vân chuyển bình quân")]
        public decimal Cvcbq { get; set; }

        [Description("Lãi vay ngân hàng")]
        public decimal Lvnh { get; set; }

        [Description("Hỗ trợ cước vân chuyển")]
        public decimal Htcvc { get; set; }

        [Description("Hỗ trợ theo văn bằng VB1370/PLX")]
        public decimal HttVb1370 { get; set; }

        [Description("Chiết Khấu Vùng 2")]
        public decimal Ckv2 { get; set; }

        [Description("CHiết khấu Xăng")]
        public decimal CkXang { get; set; }

        [Description("CHiết khấu dầu")]
        public decimal CkDau { get; set; }

        [Description("Phương thức")]
        public string PhuongThuc { get; set; }

        [Description("Thời hạn thanh toán")]
        public string Thtt { get; set; }

        [Description("Địa chỉ")]
        public string Adrress { get; set; }

        [Description("sắp xếp")]
        public int Order { get; set; }
        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblMdCustomerTnpp, CustomerTnppDto>().ReverseMap();
        }
    }
}
