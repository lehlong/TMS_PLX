using AutoMapper;
using DMS.BUSINESS.Common;
using Microsoft.EntityFrameworkCore;
using DMS.BUSINESS.Dtos.AD;
using DMS.BUSINESS.Services.AD;
using DMS.CORE.Entities.AD;
using DMS.CORE;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DMS.CORE.Entities.MD;
using Common;
using DMS.BUSINESS.Dtos.MD;

namespace DMS.BUSINESS.Services.MD
{
    public interface ITemplateListTablesDataService : IGenericService<TblMdTemplateListTablesData, TemplateListTablesDataDto>
    {
        Task<IList<TemplateListTablesDataDto>> GetAll(BaseMdFilter filter);
        //Task<IList<TemplateOpinionDataDto>> Add(IList<TemplateOpinionDataDto> dto);
        Task<Guid?> GetCode(TemplateListTablesDataGenCodeDto dto);
        Task<bool> Update(List<TemplateListTablesDataDto> dtos);

        // Phương thức xóa dựa trên orgCode, templateCode, và opinionCode
        Task Delete(TemplateListTablesDataGenCodeDto dto);
        Task<byte[]> Export(string templateCode);
    }
    public class TemplateListTablesDataService(AppDbContext dbContext, IMapper mapper) : GenericService<TblMdTemplateListTablesData, TemplateListTablesDataDto>(dbContext, mapper), ITemplateListTablesDataService
    {
        public override async Task<PagedResponseDto> Search(BaseFilter filter)
        {
            try
            {
                var query = _dbContext.tblMdTemplateListTablesData.AsQueryable();

                if (!string.IsNullOrWhiteSpace(filter.KeyWord))
                {
                    query = query.Where(x =>
                    x.TemplateCode.Contains(filter.KeyWord) || x.OrgCode.Contains(filter.KeyWord));
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
        public  async Task<IList<TemplateListTablesDataDto>> GetAll(BaseMdFilter filter)
        {
            try
            {
                var query = _dbContext.tblMdTemplateListTablesData.AsQueryable();
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

        public async Task<Guid?> GetCode(TemplateListTablesDataGenCodeDto dto)
        {
            try
            {
                var data = await _dbContext.tblMdTemplateListTablesData
                                           .Where(x => x.OrgCode == dto.OrgCode && x.TemplateCode == dto.TemplateCode && x.ListTablesCode == dto.ListTablesCode)
                                           .FirstOrDefaultAsync();
                return data?.Code;
            }
            catch (Exception ex)
            {
                Status = false;
                Exception = ex;
                return null;
            }
        }

        public async Task Delete(TemplateListTablesDataGenCodeDto dto)
        {
            try
            {
                // Lấy code dựa trên ba giá trị
                var code = await GetCode(dto);

                if (code.HasValue)
                {
                    // Gọi phương thức Delete từ GenericService để xóa bản ghi
                    await Delete(code.Value);
                }
                else
                {
                    Status = false;
                    MessageObject.Code = "0000";
                }
            }
            catch (Exception ex)
            {
                Status = false;
                Exception = ex;
            }
        }
        public async Task<bool> Update(List<TemplateListTablesDataDto> dtos)
        {
            try
            {
                if (dtos == null || !dtos.Any())
                {
                    throw new ArgumentException("Danh sách dữ liệu cập nhật không được rỗng");
                }

                var templateCode = dtos.First().TemplateCode;
                var orgCode = dtos.First().OrgCode;

                // Tìm và xóa dữ liệu hiện có
                var existingData = await _dbContext.tblMdTemplateListTablesData
                    .Where(x => x.TemplateCode == templateCode && x.OrgCode == orgCode)
                    .ToListAsync();

                _dbContext.tblMdTemplateListTablesData.RemoveRange(existingData);

                // Nếu listTablesCode không null, thêm dữ liệu mới
                if (dtos.Any(dto => dto.ListTablesCode != null))
                {
                    var newEntities = _mapper.Map<List<TblMdTemplateListTablesData>>(dtos);
                    await _dbContext.tblMdTemplateListTablesData.AddRangeAsync(newEntities);
                }
                await _dbContext.SaveChangesAsync();

                Status = true;
                return true;
            }
            catch (Exception ex)
            {
                Status = false;
                Exception = ex;
                return false;
            }
        }
        public async Task<byte[]> Export(string templateCode)
        {
            try
            {
                var template = await _dbContext.tblMdTemplateListTables
                    .FirstOrDefaultAsync(x => x.Code == templateCode);

                if (template == null)
                {
                    throw new Exception("Template không tồn tại");
                }

                var listTablesQuery = _dbContext.TblBuListTables
                    .Where(x => x.MgCode == template.MgListTablesCode)
                    .OrderBy(x => x.OrderNumber);

                var listTablesDict = await listTablesQuery.ToDictionaryAsync(x => x.Code, x => x);

                var query = _dbContext.tblMdTemplateListTablesData
                    .Where(x => x.TemplateCode == templateCode)
                    .Include(x => x.Organize)
                    .Include(x => x.ListTables)
                    .AsQueryable();

                var data = await query.Select(x => new TemplateListTablesDataListDto
                {
                    OrgCode = x.OrgCode,
                    OrgName = x.Organize.Name,
                    ListTablesId = x.ListTables.Id,
                    ListTablesName = x.ListTables.Name,
                    Unit = "đồng",
                    AuditValue = 0,
                    AuditNotes = "",
                    ExplanationValue = 0,
                    Explanation = "",
                    OrderNumber = x.ListTables.OrderNumber
                }).ToListAsync();

                var sortedData = data.GroupBy(x => x.OrgCode).SelectMany(group => group.OrderBy(x => x.OrderNumber))
                     .ToList();

                // Thêm trường IsParent vào TemplateListTablesDataListDto
                var parentItems = listTablesDict.Values.Where(x => x.PId == "LTB").Select(x => x.Id).ToHashSet();
                foreach (var item in sortedData)
                {
                    item.IsParent = parentItems.Contains(item.ListTablesId);
                }

                sortedData.Insert(0, new TemplateListTablesDataListDto
                {
                    OrgCode = "Tổng cộng",
                    OrgName = "",
                    ListTablesId = "",
                    ListTablesName = "",
                    Unit = "",
                    AuditValue = 0,
                    AuditNotes = "",
                    ExplanationValue = 0,
                    Explanation = "",
                });

                var additionalInfo = new Dictionary<string, string>
                {
                    { "Người tạo", template.OrgCode ?? "Admin" },
                    { "Người duyệt", "" }
                };

                return await ExportExtension.ExportToExcel(sortedData, "Sheet 1", $" {template.Name}", additionalInfo);
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
