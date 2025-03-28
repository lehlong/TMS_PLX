using AutoMapper;
using DMS.BUSINESS.Common;
using DMS.BUSINESS.Dtos.MD;
using DMS.CORE.Entities.MD;
using DMS.CORE;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DMS.BUSINESS.Services.BU;
using Common;
using DMS.BUSINESS.Filter.MD;

namespace DMS.BUSINESS.Services.MD
{
    public interface IMgListTablesService : IGenericService<TblMdMgListTables, MgListTablesDto>
    {
        Task<IList<MgListTablesDto>> GetAll(BaseMdFilter filter);
        Task<byte[]> Export(BaseMdFilter filter);
        Task<PagedResponseDto> Search(MgListTablesFilter filter);
    }
    public class MgListTablesService(AppDbContext dbContext, IMapper mapper) : GenericService<TblMdMgListTables, MgListTablesDto>(dbContext, mapper), IMgListTablesService
    {
        public async Task<PagedResponseDto> Search(MgListTablesFilter filter)
        {
            try
            {
                var query = _dbContext.tblMdMgListTables.AsQueryable();
                if (!string.IsNullOrWhiteSpace(filter.KeyWord))
                {
                    query = query.Where(x => x.Code.ToString().Contains(filter.KeyWord) || 
                                            x.Name.ToString().Contains(filter.KeyWord) || 
                                            x.TimeYear.ToString().Contains(filter.KeyWord));
                }
                if (!string.IsNullOrWhiteSpace(filter.GroupCode.ToString()))
                {
                    query = query.Where(x => x.GroupCode.ToString().Contains(filter.GroupCode.ToString()));

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
        public async Task<byte[]> Export(BaseMdFilter filter)
        {
            try
            {
                var query = _dbContext.tblMdMgListTables.AsQueryable();
                if (!string.IsNullOrWhiteSpace(filter.KeyWord))
                {
                    query = query.Where(x => x.AuditPeriod.Equals(filter.KeyWord));
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
        public async Task<IList<MgListTablesDto>> GetAll(BaseMdFilter filter)
        {
            try
            {
                var query = _dbContext.tblMdMgListTables.AsQueryable();
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
    }
}
