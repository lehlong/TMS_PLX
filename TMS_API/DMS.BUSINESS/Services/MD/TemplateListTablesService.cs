using AutoMapper;
using Common;
using DMS.BUSINESS.Common;
using DMS.BUSINESS.Dtos.MD;
using DMS.CORE.Entities.MD;
using DMS.CORE;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DMS.BUSINESS.Dtos.AD;
using DMS.BUSINESS.Dtos.BU;
using DMS.BUSINESS.Services.AD;
using DMS.BUSINESS.Services.BU;
using DMS.CORE.Entities.AD;
using DMS.BUSINESS.Filter.MD;

namespace DMS.BUSINESS.Services.MD
{
    public interface ITemplateListTablesService : IGenericService<TblMdTemplateListTables, TemplateListTablesDto>
    {
        Task<IList<TemplateListTablesDto>> GetAll(BaseMdFilter filter);

        Task<PagedResponseDto> Search(TemplateListTablesGroupsFilter filter);

        //Task<byte[]> Export(BaseMdFilter filter);
        Task<TemplateListTablesDetailDto> GetTemplateWithTree(string code);
        Task<bool> ChangeIsActiveStatus(string code);
    }
    public class TemplateListTablesService(AppDbContext dbContext, IMapper mapper) : GenericService<TblMdTemplateListTables, TemplateListTablesDto>(dbContext, mapper), ITemplateListTablesService
    {

        public async Task<PagedResponseDto> Search(TemplateListTablesGroupsFilter filter)
        {
            try
            {
                var query = _dbContext.tblMdTemplateListTables.AsQueryable();

                if (!string.IsNullOrWhiteSpace(filter.KeyWord))
                {
                    query = query.Where(x =>
                    x.Name.Contains(filter.KeyWord) || x.Code.Contains(filter.KeyWord));
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

        public async Task<IList<TemplateListTablesDto>> GetAll(BaseMdFilter filter)
        {
            try
            {
                var query = _dbContext.tblMdTemplateListTables.AsQueryable();
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

        public async Task<TemplateListTablesDetailDto> GetTemplateWithTree(string Code)
        {
            var data = await _dbContext.tblMdTemplateListTables.Include(x => x.TemDataReferences).FirstOrDefaultAsync(x => x.Code == Code);

            if (data == null) return null;

            var lstNodeOrg = new List<OrganizeDto>();
            var rootNodeOrg = new OrganizeDto() { Id = "ORG", PId = "-ORG", Name = "Danh Sách Các Đơn Vị", Title = "1.1_  STC", Key = "ORG" };
            lstNodeOrg.Add(rootNodeOrg);
            var lstNodeLtb = new List<ListTablesDto>();
            var rootNodeLtb = new ListTablesDto() { Id = "LTB", PId = "-", Name = "Danh Mục Chỉ Tiêu", Title = "LTB. DANH MỤC CHỈ TIÊU", Key = "LTB" };
            lstNodeLtb.Add(rootNodeLtb);
            var lstAllListTables = await _dbContext.TblBuListTables
                                                   .Where(x => x.MgCode == data.MgListTablesCode)
                                                   .OrderBy(x => x.OrderNumber)
                                                   .ToListAsync();
            var lstAllOrganize = await _dbContext.tblAdOrganize
                                                    .Where(x => x.IsActive == true)
                                                    .OrderBy(x => x.OrderNumber)
                                                    .ToListAsync();
            var lstListTablesInTemplate = data.TemDataReferences
                .Select(x => x.ListTablesCode)
                .ToList();
            var lstOrganizeInTemplate = data.TemDataReferences
                .Select(x => x.OrgCode)
                .ToList();

            foreach (var tem in lstAllOrganize)
            {
                var nodeOrg = new OrganizeDto()
                {
                    Id = tem.Id,
                    Name = tem.Name,
                    PId = tem.PId,
                    OrderNumber = tem.OrderNumber,
                    Title = $"{tem.Id}_{tem.Name}",
                    Key = tem.Id,
                    IsActive = tem.IsActive
                };
                if (lstOrganizeInTemplate.Contains(tem.Id))
                {
                    nodeOrg.IsChecked = true;
                }
                lstNodeOrg.Add(nodeOrg);
            }

            var nodeDict = lstNodeOrg.ToDictionary(n => n.Id);
            foreach (var item in lstNodeOrg)
            {
                if (item.PId == "-ORG" || !nodeDict.TryGetValue(item.PId, out OrganizeDto parentNode))
                {
                    continue;
                }

                parentNode.Children ??= new List<OrganizeDto>();
                parentNode.Children.Add(item);
            }
            foreach (var tem in lstAllListTables)
            {
                var nodeOpi = new ListTablesDto()
                {
                    Code = tem.Code,
                    Id = tem.Id,
                    Name = tem.Name,
                    PId = tem.PId,
                    OrderNumber = tem.OrderNumber,
                    Title = $"{tem.Id} _ {tem.Name}",
                    Key = tem.Id,
                    MgCode = tem.MgCode
                };
                if (lstListTablesInTemplate.Contains(tem.Code) && lstOrganizeInTemplate.Contains(tem.PId))
                {
                    nodeOpi.IsChecked = true;
                }
                lstNodeLtb.Add(nodeOpi);
            }

            var nodeDictOpi = lstNodeLtb.ToDictionary(n => n.Id);
            foreach (var item in lstNodeLtb)
            {
                if (item.PId == "-" || !nodeDictOpi.TryGetValue(item.PId, out ListTablesDto parentNode))
                {
                    continue;
                }

                parentNode.Children ??= new List<ListTablesDto>();
                parentNode.Children.Add(item);
            }
            foreach (var organize in lstNodeOrg)
            {
                if (organize.Children != null)
                {
                    foreach (var child in organize.Children)
                    {
                        var flatListTables = lstNodeLtb.Select(opi => new ListTablesDto
                        {
                            Code = opi.Code,
                            Id = opi.Id,
                            Name = opi.Name,
                            PId = opi.PId,
                            OrderNumber = opi.OrderNumber,
                            Title = opi.Title,
                            Key = opi.Key,
                            MgCode = opi.MgCode,
                            IsChecked = data.TemDataReferences.Any(x => x.OrgCode == child.Id && x.ListTablesCode == opi.Code)
                        }).ToList();
                        child.TreeListTables = new List<ListTablesDto> { BuildListTablesTree(flatListTables) };

                    }
                }
            }


            var result = _mapper.Map<TemplateListTablesDetailDto>(data);
            result.TreeOrganize = rootNodeOrg;
            result.TreeListTables = rootNodeLtb;
            return result;
        }
        private ListTablesDto BuildListTablesTree(List<ListTablesDto> flatList, string parentId = "LTB")
        {
            var root = flatList.FirstOrDefault(x => x.Id == parentId);
            if (root == null) return null;

            root.Children = flatList
                .Where(x => x.PId == root.Id)
                .Select(child => BuildListTablesTree(flatList, child.Id))
                .ToList();

            return root;
        }



        public async Task<bool> ChangeIsActiveStatus(string code)
        {
            try
            {
                var temIsActive = await _dbContext.tblMdTemplateListTables.FindAsync(code);

                if (temIsActive == null)
                {
                    return false;
                }
                var mgListTables = _dbContext.tblMdMgListTables.FirstOrDefault(x => x.Code == temIsActive.MgListTablesCode);
                // Đảo ngược trạng thái IsClosed
                temIsActive.IsActive = !temIsActive.IsActive;
                mgListTables.IsActive = !mgListTables.IsActive;
                await _dbContext.SaveChangesAsync();

                return true;
            }
            catch (Exception ex)
            {
                Status = false;
                Exception = ex;
                return false;
            }
        }
    }

}
