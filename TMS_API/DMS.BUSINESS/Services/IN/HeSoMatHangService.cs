using AutoMapper;
using Common;
using DMS.BUSINESS.Common;
using DMS.BUSINESS.Dtos.IN;
using DMS.CORE;
using DMS.CORE.Entities.IN;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMS.BUSINESS.Services.IN
{
    public interface IHeSoMatHangService : IGenericService<TblInHeSoMatHang, HeSoMatHangDto>
    {
        Task<IList<HeSoMatHangDto>> GetAll(BaseMdFilter filter);
        Task AddHeSoMatHang(HeSoMatHangDto heSoMatHang);
        Task<byte[]> Export(BaseMdFilter filter);
    }
    public class HeSoMatHangService(AppDbContext dbContext, IMapper mapper) : GenericService<TblInHeSoMatHang, HeSoMatHangDto>(dbContext, mapper), IHeSoMatHangService
    {
        public override async Task<PagedResponseDto> Search(BaseFilter filter)
        {
            try
            {
                var query = _dbContext.TblInHeSoMatHang.AsQueryable();

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
        public async Task<IList<HeSoMatHangDto>> GetAll(BaseMdFilter filter)
        {
            try
            {
                var query = _dbContext.TblInHeSoMatHang.AsQueryable();
                if (filter.IsActive.HasValue)
                {
                    query = query.Where(x => x.IsActive == filter.IsActive);
                }
                return await base.GetAllMd(query, filter);
            }
            catch (Exception ex)
            {
                Status = false;
                Exception = ex;
                return null;
            }
        }


        public async Task AddHeSoMatHang(HeSoMatHangDto heSoMatHang)
        {
            try
            {
                var lstHeSoMatHang = new TblInHeSoMatHang()
                {
                    Code = heSoMatHang.GoodsCode + "-" + DateTime.Now.ToString("ddMMyyyy"),
                    GoodsCode = heSoMatHang.GoodsCode,
                    HeSoVcf = heSoMatHang.HeSoVcf,
                    ThueBvmt = heSoMatHang.ThueBvmt,
                    L15ChuaVatBvmt = heSoMatHang.L15ChuaVatBvmt,
                    FromDate = heSoMatHang.FromDate,
                    ToDate = heSoMatHang.ToDate,
                    GiamGiaFob = heSoMatHang.GiamGiaFob,
                    IsActive = heSoMatHang.IsActive,
                };
                this._dbContext.TblInHeSoMatHang.Add(lstHeSoMatHang);
                await this._dbContext.SaveChangesAsync();
            }
            catch (Exception ex)
            {
                Status = false;
                Exception = ex;
            }
        }

        public async Task<byte[]> Export(BaseMdFilter filter)
        {
            try
            {
                var query = _dbContext.TblInHeSoMatHang.AsQueryable();
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
