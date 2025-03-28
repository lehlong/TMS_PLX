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
    public interface IVinhCuaLoService : IGenericService<TblInVinhCuaLo, VinhCuaLoDto>
    {
        Task<IList<VinhCuaLoDto>> GetAll(BaseMdFilter filter);
        Task AddVinhCuaLo(VinhCuaLoDto vinhCuaLo);
        Task UpdateVinhCuaLo(VinhCuaLoDto vinhCuaLo);
        Task<byte[]> Export(BaseMdFilter filter);
    }
    public class VinhCuaLoService(AppDbContext dbContext, IMapper mapper) : GenericService<TblInVinhCuaLo, VinhCuaLoDto>(dbContext, mapper), IVinhCuaLoService
    {
        public override async Task<PagedResponseDto> Search(BaseFilter filter)
        {
            try
            {
                var query = _dbContext.TblInVinhCuaLo.AsQueryable();

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
        public async Task<IList<VinhCuaLoDto>> GetAll(BaseMdFilter filter)
        {
            try
            {
                var query = _dbContext.TblInVinhCuaLo.AsQueryable();
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

        public async Task AddVinhCuaLo(VinhCuaLoDto vinhCuaLo)
        {
            try
            {
                var lstVinhCuaLo = new TblInVinhCuaLo()
                {
                    Code = vinhCuaLo.GoodsCode + "-" + DateTime.Now.ToString("ddMMyyyy"),
                    GoodsCode = vinhCuaLo.GoodsCode,
                    GblcsV1 = vinhCuaLo.GblcsV1,
                    GblV2 = vinhCuaLo.GblV2,
                    V2_V1 = vinhCuaLo.GblV2 - vinhCuaLo.GblcsV1,
                    MtsV1 = vinhCuaLo.MtsV1,
                    Gny = vinhCuaLo.GblcsV1 + vinhCuaLo.MtsV1,
                    Clgblv = vinhCuaLo.GblV2 - (vinhCuaLo.GblcsV1 + vinhCuaLo.MtsV1),
                    FromDate = vinhCuaLo.FromDate,
                    ToDate = vinhCuaLo.ToDate,
                    IsActive = vinhCuaLo.IsActive,
                };
                this._dbContext.TblInVinhCuaLo.Add(lstVinhCuaLo);
                await this._dbContext.SaveChangesAsync();
            }
            catch (Exception ex)
            {
                Status = false;
                Exception = ex;
            }
        }

        public async Task UpdateVinhCuaLo(VinhCuaLoDto vinhCuaLo)
        {
            try
            {
                var lstVinhCuaLo = new TblInVinhCuaLo()
                {
                    Code = vinhCuaLo.Code,
                    GoodsCode = vinhCuaLo.GoodsCode,
                    GblcsV1 = vinhCuaLo.GblcsV1,
                    GblV2 = vinhCuaLo.GblV2,
                    V2_V1 = vinhCuaLo.GblV2 - vinhCuaLo.GblcsV1,
                    MtsV1 = vinhCuaLo.MtsV1,
                    Gny = vinhCuaLo.GblcsV1 + vinhCuaLo.MtsV1,
                    Clgblv = vinhCuaLo.GblV2 - (vinhCuaLo.GblcsV1 + vinhCuaLo.MtsV1),
                    FromDate = vinhCuaLo.FromDate,
                    ToDate = vinhCuaLo.ToDate,
                    IsActive = vinhCuaLo.IsActive,
                };
                this._dbContext.TblInVinhCuaLo.Update(lstVinhCuaLo);
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
                var query = _dbContext.TblInVinhCuaLo.AsQueryable();
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
