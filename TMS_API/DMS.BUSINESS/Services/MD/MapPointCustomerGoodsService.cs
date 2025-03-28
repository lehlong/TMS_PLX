using AutoMapper;
using Common;
using DMS.BUSINESS.Common;
using DMS.BUSINESS.Dtos.MD;
using DMS.BUSINESS.Filter.MD;
using DMS.CORE;
using DMS.CORE.Entities.MD;
using Microsoft.AspNetCore.Mvc.Filters;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMS.BUSINESS.Services.MD
{
    public interface IMapPointCustomerGoodsService : IGenericService<TblMdMapPointCustomerGoods, MapPointCustomerGoodsDto>
    {
        Task<IList<MapPointCustomerGoodsDto>> GetAll(BaseMdFilter filter);
        Task<PagedResponseDto> Search(MappingCustomerPointFilter filter);
        Task<byte[]> Export(BaseMdFilter filter);
    }
    public class MapPointCustomerGoodsService(AppDbContext dbContext, IMapper mapper) : GenericService<TblMdMapPointCustomerGoods, MapPointCustomerGoodsDto>(dbContext, mapper), IMapPointCustomerGoodsService
    {
        public async Task<PagedResponseDto> Search(MappingCustomerPointFilter filter)
        {
            try
            {
                var query = _dbContext.TblMdMapPointCustomerGoods.AsQueryable();

                if (!string.IsNullOrWhiteSpace(filter.KeyWord))
                {
                    query = query.Where(x =>
                    x.Code.Contains(filter.KeyWord));
                }
                if (!string.IsNullOrEmpty(filter.CustomerCode))
                {
                    query = query.Where(x =>
                    x.CustomerCode.Contains(filter.CustomerCode));
                }
                if (!string.IsNullOrEmpty(filter.PointCode))
                {
                    query = query.Where(x =>
                    x.DeliveryPointCode.Contains(filter.PointCode));
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
        public async Task<IList<MapPointCustomerGoodsDto>> GetAll(BaseMdFilter filter)
        {
            try
            {
                var query = _dbContext.TblMdMapPointCustomerGoods.AsQueryable();
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
        public async Task<byte[]> Export(BaseMdFilter filter)
        {
            try
            {
                var query = _dbContext.TblMdMapPointCustomerGoods.AsQueryable();
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
