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
using Common;

namespace DMS.BUSINESS.Services.MD
{
    public interface IMgListTablesGroupsService : IGenericService<TblMdMgListTablesGroups, MgListTablesGroupsDto>
    {
        Task<IList<MgListTablesGroupsDto>> GetAll(BaseMdFilter filter);
    }
    public class MgListTablesGroupsService(AppDbContext dbContext, IMapper mapper) : GenericService<TblMdMgListTablesGroups, MgListTablesGroupsDto>(dbContext, mapper), IMgListTablesGroupsService
    {
        public override async Task<PagedResponseDto> Search(BaseFilter filter)
        {
            try
            {
                var query = _dbContext.tblMdMgListTablesGroup.AsQueryable();
                if (!string.IsNullOrWhiteSpace(filter.KeyWord))
                {
                    query = query.Where(x => x.Code.ToString().Contains(filter.KeyWord) || x.Id.ToString().Contains(filter.KeyWord) ||
                                            x.Name.ToString().Contains(filter.KeyWord));
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
        public async Task<IList<MgListTablesGroupsDto>> GetAll(BaseMdFilter filter)
        {
            try
            {
                var query = _dbContext.tblMdMgListTablesGroup.AsQueryable();
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
