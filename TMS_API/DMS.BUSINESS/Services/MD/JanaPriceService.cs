using AutoMapper;
using Common;
using DMS.BUSINESS.Common;
using DMS.BUSINESS.Dtos.MD;
using DMS.CORE;
using DMS.CORE.Common;
using DMS.CORE.Entities.MD;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static DMS.BUSINESS.Services.MD.JanaPriceService;

namespace DMS.BUSINESS.Services.MD
{
    public interface IJanaPriceService : IGenericService<TblMdJanaPrice, JanaPriceDto>
    {
        Task<IList<JanaPrice>> GetAll(BaseMdFilter filter);
        Task<byte[]> Export(BaseMdFilter filter);
    }
    public class JanaPriceService(AppDbContext dbContext, IMapper mapper) : GenericService<TblMdJanaPrice, JanaPriceDto>(dbContext, mapper), IJanaPriceService
    {
        public override async Task<PagedResponseDto> Search(BaseFilter filter)
        {
            try
            {
                var query = _dbContext.TblMdJanaPrice.AsQueryable();

                if (!string.IsNullOrWhiteSpace(filter.KeyWord))
                {
                    query = query.Where(x =>
                    x.Code.Contains(filter.KeyWord));
                }
                if (filter.IsActive.HasValue)
                {
                    query = query.Where(x => x.IsActive == filter.IsActive);
                }
                return await Paging(query, filter);
            }
            catch (Exception ex)
            {
                Status = false;
                Exception = ex;
                return null;
            }
        }
        public class JanaPrice 
        {
            public string Code { get; set; }
            public string PtppCode { get; set; }

            public string GoodsCode { get; set; }
            public string GoodsName { get; set; }

            public decimal GiaNhapMuaVat { get; set; }

          
            public decimal LuongNxNoiBo { get; set; }

           
            public decimal CpHoTro { get; set; }

        
            public decimal GiaDaiLyVat { get; set; }

           
            public decimal LuongBanHang { get; set; }

        }
        public async Task<IList<JanaPrice>> GetAll(BaseMdFilter filter)
        {
            try
            {
                var queryJaPrice = _dbContext.TblMdJanaPrice.AsQueryable();
                var querygoods = _dbContext.TblMdJaGoods.AsQueryable();
                var queryPtpp = _dbContext.TblMdJaPtPhanPhoi.AsQueryable();
                var lstJaprice = new List<JanaPrice>();
                foreach (var good in querygoods)
                {
                    foreach (var pptt in queryPtpp)
                    {
                        var DK = queryJaPrice.FirstOrDefault(x => x.GoodsCode == good.Code && x.PtppCode == pptt.Code);
                        var item = new JanaPrice()
                        {
                            GoodsName = good.Name,
                            GoodsCode = good.Code,
                            PtppCode = pptt.Code,
                            GiaNhapMuaVat = DK?.GiaNhapMuaVat??0 ,
                            LuongBanHang=DK?.LuongBanHang??0,
                            GiaDaiLyVat=DK?.GiaDaiLyVat ?? 0,
                            LuongNxNoiBo=DK?.LuongNxNoiBo ?? 0,
                            CpHoTro=DK?.CpHoTro ?? 0

                        };
                        lstJaprice.Add(item);

                    }
                }

                return lstJaprice;
            }
            catch (Exception ex)
            {
                Status = false;
                Exception = ex;
                return null;
            }
        }
        public async Task<byte[]> Export(BaseMdFilter filter)
        {
            try
            {
                var query = _dbContext.TblMdJanaPrice.AsQueryable();
                if (!string.IsNullOrWhiteSpace(filter.KeyWord))
                {
                    query = query.Where(x => x.Code.Contains(filter.KeyWord));
                }
                if (filter.IsActive.HasValue)
                {
                    query = query.Where(x => x.IsActive == filter.IsActive);
                }
                var data = await base.GetAllMd(query, filter);
                int i = 1;
                data.ForEach(x =>
                {
                    x.OrdinalNumber = i++;
                });
                return await ExportExtension.ExportToExcel(data);
            }
            catch (Exception ex)
            {
                Status = false;
                Exception = ex;
                return null;
            }
        }

    }
}
