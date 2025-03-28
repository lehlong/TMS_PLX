using AutoMapper;
using DMS.BUSINESS.Common;
using DMS.BUSINESS.Dtos.BU;
using DMS.BUSINESS.Services.BU;
using DMS.CORE.Entities.BU;
using DMS.CORE;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DMS.CORE.Entities.MD;
using DMS.BUSINESS.Dtos.MD;
using Common;
using Microsoft.EntityFrameworkCore;
using DMS.BUSINESS.Common.Enum;
using Microsoft.AspNetCore.Http;
using OfficeOpenXml;
using static System.Runtime.CompilerServices.RuntimeHelpers;

namespace DMS.BUSINESS.Services.MD
{
    public interface IAuditPeriodListTablesService : IGenericService<TblMdAuditPeriodListTables, AuditPeriodListTablesDto>
    {
        Task<IList<AuditPeriodListTablesDto>> GetAll(BaseMdFilter filter);
        Task<IList<MgListTablesDto>> GetMgListTablesByAuditPeriodCode(string auditPeriodCode);
        Task<IList<TemplateListTablesDto>> GetTemplateListTablesByAuditPeriodCode(string auditPeriodCode, Guid templateGroupCode);
        Task<IList<AuditTemplateListTablesDataListDto>> GetTemDataWithMgListTables(int auditListTablesCode);
        Task<IList<AuditTemplateListTablesDataListDto>> GetTemDataWithMgListTablesAndOrgCode(int auditListTablesCode, string? OrgCode);
        Task<AuditPeriodListTablesDto> Add(AuditPeriodListTablesCreateDto dto);
        Task UpdateAuditTemplateListTablesData(AuditTemplateListTablesDataDto dto);
        Task<bool> ImportExcelAndUpdateDataSTC(IFormFile file, int auditListTablesCode);
        Task<bool> ImportExcelAndUpdateDataDV(IFormFile file, int auditListTablesCode);
        Task Delete(int auditListTablesCode);
        Task<bool> ChangeStatusReview(int code, string? textContent);
        Task<bool> ChangeStatusApproval(int code);
        Task<bool> ChangeStatusCancel(int code, string? action, string? textContent);
        Task<bool> ChangeStatusconfirm(int code, string? action, string? textContent);
        Task<byte[]> ExportAuditTemplateListTablesData(int auditListTablesCode, string? orgCode);
        Task<IList<AuditTemplateHistoryDto>> GetHistoryByListAuditCode(string listAuditCode);
    }
    public class AuditPeriodListTablesService(AppDbContext dbContext, IMapper mapper) : GenericService<TblMdAuditPeriodListTables, AuditPeriodListTablesDto>(dbContext, mapper), IAuditPeriodListTablesService
    {
        public override async Task<PagedResponseDto> Search(BaseFilter filter)
        {
            try
            {
                var query = _dbContext.tblMdAuditPeriodListTables.AsQueryable();
                if (!string.IsNullOrWhiteSpace(filter.KeyWord))
                {
                    query = query.Where(x => x.Code.ToString().Contains(filter.KeyWord) ||
                                            x.AuditPeriodCode.ToString().Contains(filter.KeyWord) ||
                                            x.ListTablesCode.ToString().Contains(filter.KeyWord));
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
        public async Task<IList<AuditPeriodListTablesDto>> GetAll(BaseMdFilter filter)
        {
            try
            {
                var query = _dbContext.tblMdAuditPeriodListTables.AsQueryable();
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
        public async Task<IList<AuditTemplateHistoryDto>> GetHistoryByListAuditCode(string listAuditCode)
        {
            try
            {
                var history = await _dbContext.tblMdAuditTemplateHistory
                    .Where(h => h.ListAuditCode == listAuditCode)
                    .OrderByDescending(h => h.CreateDate)
                    .ToListAsync();

                return _mapper.Map<IList<AuditTemplateHistoryDto>>(history);
            }
            catch (Exception ex)
            {
                Status = false;
                Exception = ex;
                return null;
            }
        }
        public async Task UpdateAuditTemplateListTablesData(AuditTemplateListTablesDataDto dto)
        {
            try
            {
                var existingEntity = await _dbContext.tblMdAuditTemplateListTablesData
                    .FirstOrDefaultAsync(td => td.TemplateDataCode == dto.TemplateDataCode &&
                                                td.AuditListTablesCode == dto.AuditListTablesCode);

                if (existingEntity != null)
                {
                    // Cập nhật thực thể hiện có
                    _mapper.Map(dto, existingEntity);
                }
                else
                {
                    // Thêm mới nếu không tìm thấy
                    var newEntity = _mapper.Map<TblMdAuditTemplateListTablesData>(dto);
                    await _dbContext.tblMdAuditTemplateListTablesData.AddAsync(newEntity);
                }

                await _dbContext.SaveChangesAsync();

            }
            catch (Exception ex)
            {
                Status = false;
                Exception = ex;
            }
        }
        public async Task<IList<MgListTablesDto>> GetMgListTablesByAuditPeriodCode(string auditPeriodCode)
        {
            var auditPeriodData = await _dbContext.tblMdAuditPeriodListTables
                .Where(x => x.AuditPeriodCode == auditPeriodCode)
                .Include(x => x.TblMdTemplateListTables)
                .ToListAsync();
            if (auditPeriodData == null)
            {
                return null;
            }
            var mgListTables = await _dbContext.tblMdMgListTables.Where(x => x.IsActive == true).ToListAsync();

            var result = mgListTables.Select(mgListTable => new MgListTablesDto
            {
                Code = mgListTable.Code,
                Name = mgListTable.Name,
                Description = mgListTable.Description,
                TimeYear = mgListTable.TimeYear,
                AuditPeriod = mgListTable.AuditPeriod,
                IsChecked = auditPeriodData.Any(a => a.TblMdTemplateListTables.Code == mgListTable.Code)
            }).ToList();

            return result;
        }
        public async Task<IList<TemplateListTablesDto>> GetTemplateListTablesByAuditPeriodCode(string auditPeriodCode, Guid templateGroupCode)
        {
            var auditPeriodData = await _dbContext.tblMdAuditPeriodListTables
                .Where(x => x.AuditPeriodCode == auditPeriodCode)
                .Include(x => x.TblMdTemplateListTables)
                .ToListAsync();
            if (auditPeriodData == null)
            {
                return null;
            }
            var mgListTables = await _dbContext.tblMdTemplateListTables.Where(x => x.IsActive == true && x.GroupCode == templateGroupCode).ToListAsync();

            var result = mgListTables.Select(mgListTable => new TemplateListTablesDto
            {
                Code = mgListTable?.Code,
                Name = mgListTable?.Name,
                OrgCode = mgListTable?.OrgCode,
                TimeYear = mgListTable?.TimeYear,
                MgListTablesCode = mgListTable?.MgListTablesCode,
                Note = mgListTable?.Note,
                IsChecked = auditPeriodData.Any(a => a.TblMdTemplateListTables.Code == mgListTable.Code),
                GroupCode = mgListTable.GroupCode,
                AuditPeriodCode = auditPeriodData.FirstOrDefault(a => a.TblMdTemplateListTables.Code == mgListTable.Code)?.Code
            }).ToList();

            return result;
        }
        public async Task<AuditPeriodListTablesDto> Add(AuditPeriodListTablesCreateDto dto)
        {
            try
            {
                var auditPeriodListTables = _mapper.Map<TblMdAuditPeriodListTables>(dto);
                await _dbContext.tblMdAuditPeriodListTables.AddAsync(auditPeriodListTables);
                await _dbContext.SaveChangesAsync();

                int auditListTablesCode = auditPeriodListTables.Code;
                var templateListTables = await _dbContext.tblMdTemplateListTables
                    .Where(t => t.Code == auditPeriodListTables.ListTablesCode)
                    .FirstOrDefaultAsync();
                var listTables = await _dbContext.TblBuListTables
                    .Where(t => t.MgCode == templateListTables.MgListTablesCode)
                    .FirstOrDefaultAsync();
                // Get corresponding TemplateListTablesData
                var templateListTablesData = await _dbContext.tblMdTemplateListTablesData
                    .Where(t => t.TemplateCode == templateListTables.Code)
                    .ToListAsync();
                var auditTemplateListTablesDataList = templateListTablesData.Select(templateData => new TblMdAuditTemplateListTablesData
                {
                    TemplateDataCode = templateData.Code,
                    AuditListTablesCode = auditListTablesCode, // Sử dụng Code vừa mới được inser
                    Unit = listTables.CurrencyCode.ToString(),
                    AuditValue = null,
                    AuditExplanation = null,
                    ExplanationValue = null,
                    ExplanationNote = null
                }).ToList();

                // Thêm vào bảng TblMdAuditTemplateListTablesData
                await _dbContext.tblMdAuditTemplateListTablesData.AddRangeAsync(auditTemplateListTablesDataList);
                await _dbContext.SaveChangesAsync();

                return _mapper.Map<AuditPeriodListTablesDto>(auditPeriodListTables);
            }
            catch (Exception ex)
            {
                Status = false;
                Exception = ex;
                return null;
            }
        }

        public async Task<IList<AuditTemplateListTablesDataListDto>> GetTemDataWithMgListTables(int auditListTablesCode)
        {
            try
            {
                var auditPeriodListTables = await _dbContext.tblMdAuditPeriodListTables
                 .Include(x => x.TblMdTemplateListTables)
                 .FirstOrDefaultAsync(x => x.Code == auditListTablesCode);

                if (auditPeriodListTables == null)
                {
                    return null;
                }

                var combinedData = await _dbContext.tblMdAuditTemplateListTablesData
                    .Include(a => a.templateListTablesData)
                    .Where(a => a.AuditListTablesCode == auditListTablesCode)
                    .Select(a => new AuditTemplateListTablesDataListDto
                    {
                        TemplateDataCode = a.TemplateDataCode,
                        AuditListTablesCode = a.AuditListTablesCode,
                        OrgCode = a.templateListTablesData.OrgCode,
                        ListTablesCode = a.templateListTablesData.ListTablesCode,
                        TemplateCode = a.templateListTablesData.TemplateCode,
                        Unit = a.Unit,
                        AuditValue = a.AuditValue,
                        AuditExplanation = a.AuditExplanation,
                        ExplanationValue = a.ExplanationValue,
                        ExplanationNote = a.ExplanationNote
                    })
                    .ToListAsync();

                return combinedData;
            }
            catch (Exception ex)
            {
                Status = false;
                Exception = ex;
                return null;
            }
        }
        public async Task Delete(int auditPeriodCode)
        {
            try
            {
                await _dbContext.Database.BeginTransactionAsync();

                // Find the AuditPeriodListTables entity
                var entity = await _dbContext.tblMdAuditPeriodListTables.FindAsync(auditPeriodCode);
                if (entity == null)
                {
                    Status = false;
                    MessageObject.Code = "0000";
                    return;
                }

                // Delete related AuditTemplateListTablesData
                var relatedData = await _dbContext.tblMdAuditTemplateListTablesData
                    .Where(x => x.AuditListTablesCode == auditPeriodCode)
                    .ToListAsync();
                if (relatedData == null)
                {
                    Status = false;
                    MessageObject.Code = "0000";
                    return;
                }
                _dbContext.tblMdAuditTemplateListTablesData.RemoveRange(relatedData);

                // Delete the AuditPeriodListTables entity
                _dbContext.tblMdAuditPeriodListTables.Remove(entity);

                await _dbContext.SaveChangesAsync();
                await _dbContext.Database.CommitTransactionAsync();
            }
            catch (Exception ex)
            {
                await _dbContext.Database.RollbackTransactionAsync();
                Status = false;
                Exception = ex;
            }
        }
        public async Task<bool> ChangeStatusCancel(int code, string? action, string? textContent)
        {
            try
            {
                var auditPeriodListTables = await _dbContext.tblMdAuditPeriodListTables.FindAsync(code);
                if (auditPeriodListTables == null)
                {
                    return false;
                }
                auditPeriodListTables.Status = ((int)AuditPeriodListTables.KHOI_TAO).ToString();

                await _dbContext.SaveChangesAsync();
                var h = new TblMdAuditTemplateHistory()
                {
                    Id = Guid.NewGuid().ToString(),
                    ListAuditCode = auditPeriodListTables.AuditPeriodCode,
                    Action = action,
                    TextContent = textContent,
                    AuditPeriodListTablesCode = code
                };
                await this._dbContext.tblMdAuditTemplateHistory.AddAsync(h);
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
        public async Task<bool> ChangeStatusReview(int code, string? textContent)
        {
            try
            {
                var auditPeriodListTables = await _dbContext.tblMdAuditPeriodListTables.FindAsync(code);
                if (auditPeriodListTables == null)
                {
                    return false;
                }
                auditPeriodListTables.Status = ((int)AuditPeriodListTables.CHO_PHE_DUYET).ToString();
                var h = new TblMdAuditTemplateHistory()
                {
                    Id = Guid.NewGuid().ToString(),
                    ListAuditCode = auditPeriodListTables.AuditPeriodCode,
                    Action = "Trình duyệt",
                    TextContent = textContent,
                    AuditPeriodListTablesCode = code
                };
                await this._dbContext.tblMdAuditTemplateHistory.AddAsync(h);
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

        public async Task<bool> ChangeStatusApproval(int code)
        {
            try
            {
                var auditPeriodListTables = await _dbContext.tblMdAuditPeriodListTables.FindAsync(code);
                if (auditPeriodListTables == null)
                {
                    return false;
                }
                auditPeriodListTables.Status = ((int)AuditPeriodListTables.DA_PHE_DUYET).ToString();
                var h = new TblMdAuditTemplateHistory()
                {
                    Id = Guid.NewGuid().ToString(),
                    ListAuditCode = auditPeriodListTables.AuditPeriodCode,
                    Action = "Phê duyệt",
                    TextContent = null,
                    AuditPeriodListTablesCode = code
                };
                await this._dbContext.tblMdAuditTemplateHistory.AddAsync(h);
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

        
        public async Task<bool> ChangeStatusconfirm(int code, string? action, string? textContent)
        {
            try
            {
                var auditPeriodListTables = await _dbContext.tblMdAuditPeriodListTables.FindAsync(code);
                if (auditPeriodListTables == null)
                {
                    return false;
                }
                auditPeriodListTables.Status = ((int)AuditPeriodListTables.XAC_NHAN).ToString();
                var h = new TblMdAuditTemplateHistory()
                {
                    Id = Guid.NewGuid().ToString(),
                    ListAuditCode = auditPeriodListTables.AuditPeriodCode,
                    Action = action,
                    TextContent = textContent,
                    AuditPeriodListTablesCode = code
                };
                await this._dbContext.tblMdAuditTemplateHistory.AddAsync(h);
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
        public async Task<IList<AuditTemplateListTablesDataListDto>> GetTemDataWithMgListTablesAndOrgCode(int auditListTablesCode, string? OrgCode)
        {
            try
            {
                var auditPeriodListTables = await _dbContext.tblMdAuditPeriodListTables
           .Include(x => x.TblMdTemplateListTables)
           .FirstOrDefaultAsync(x => x.Code == auditListTablesCode);

                if (auditPeriodListTables == null)
                {
                    return null;
                }

                var query = _dbContext.tblMdAuditTemplateListTablesData
                    .Include(a => a.templateListTablesData)
                    .Where(a => a.AuditListTablesCode == auditListTablesCode);

                // Apply OrgCode filter only if it's provided
                if (!string.IsNullOrWhiteSpace(OrgCode))
                {
                    query = query.Where(a => a.templateListTablesData.OrgCode == OrgCode);
                }

                var combinedData = await query
                    .Select(a => new AuditTemplateListTablesDataListDto
                    {
                        TemplateDataCode = a.TemplateDataCode,
                        AuditListTablesCode = a.AuditListTablesCode,
                        OrgCode = a.templateListTablesData.OrgCode,
                        ListTablesCode = a.templateListTablesData.ListTablesCode,
                        TemplateCode = a.templateListTablesData.TemplateCode,
                        Unit = a.Unit,
                        AuditValue = a.AuditValue,
                        AuditExplanation = a.AuditExplanation,
                        ExplanationValue = a.ExplanationValue,
                        ExplanationNote = a.ExplanationNote
                    })
                    .ToListAsync();

                return combinedData;
            }
            catch (Exception ex)
            {
                Status = false;
                Exception = ex;
                return null;
            }
        }

        public async Task<bool> ImportExcelAndUpdateDataSTC(IFormFile file, int auditListTablesCode)
        {

            if (file == null || file.Length == 0)
            {
                throw new ArgumentException("File is empty or null.");
            }

            try
            {
                var auditPeriod = await _dbContext.tblMdAuditPeriodListTables
                    .FirstOrDefaultAsync(x => x.Code == auditListTablesCode);

                if (auditPeriod == null)
                {
                    throw new InvalidOperationException($"Audit period with code {auditListTablesCode} not found.");
                }

                var updatedData = new List<AuditTemplateListTablesDataListDto>();
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var stream = new MemoryStream())
                {
                    await file.CopyToAsync(stream);
                    using (var package = new ExcelPackage(stream))
                    {
                        var worksheet = package.Workbook.Worksheets[0];
                        int rowCount = worksheet.Dimension.Rows;

                        for (int row = 9; row <= rowCount; row++)
                        {
                            var orgCodeData = worksheet.Cells[row, 1].Value?.ToString();
                            var listTablesId = worksheet.Cells[row, 3].Value?.ToString();
                            var listTablesName = worksheet.Cells[row, 4].Value?.ToString();

                            if (string.IsNullOrWhiteSpace(orgCodeData) || string.IsNullOrWhiteSpace(listTablesId) || string.IsNullOrWhiteSpace(listTablesName))
                            {
                                continue; // Skip rows with missing essential data
                            }
                            var templateList = await _dbContext.tblMdTemplateListTables
                               .FirstOrDefaultAsync(tem => tem.Code == auditPeriod.ListTablesCode);

                            var temCode = templateList?.Code;
                            var listTable = await _dbContext.TblBuListTables
                                            .FirstOrDefaultAsync(lt => lt.Id == listTablesId &&
                                            lt.Name == listTablesName &&
                                            lt.MgCode == auditPeriod.ListTablesCode);
                            var templateData = await _dbContext.tblMdTemplateListTablesData
                                    .FirstOrDefaultAsync(td => td.TemplateCode == temCode && td.ListTablesCode == listTable.Code &&
                                                    td.OrgCode == orgCodeData);
                            if (listTable == null)
                            {
                                throw new InvalidOperationException($"List table not found for Id: {listTablesId}, Name: {listTablesName}");
                            }

                            var dto = new AuditTemplateListTablesDataListDto
                            {
                                TemplateDataCode = templateData.Code,
                                AuditListTablesCode = auditListTablesCode,
                                OrgCode = orgCodeData,
                                ListTablesCode = listTable.Code,
                                Unit = worksheet.Cells[row, 5].Value?.ToString(),
                                AuditValue = decimal.TryParse(worksheet.Cells[row, 6].Value?.ToString(), out var auditValue) ? auditValue : null,
                                AuditExplanation = worksheet.Cells[row, 7].Value?.ToString(),
                              
                            };

                            updatedData.Add(dto);
                        }
                    }
                }

                // Update or insert data
                foreach (var dto in updatedData)
                {
                    var existingEntity = await _dbContext.tblMdAuditTemplateListTablesData
                        .FirstOrDefaultAsync(td => td.AuditListTablesCode == dto.AuditListTablesCode &&
                                                   td.TemplateDataCode == dto.TemplateDataCode);

                    if (existingEntity != null)
                    {
                        // Update existing entity
                        _mapper.Map(dto, existingEntity);
                    }
                    else
                    {
                        // Insert new entity
                        var newEntity = _mapper.Map<TblMdAuditTemplateListTablesData>(dto);
                        await _dbContext.tblMdAuditTemplateListTablesData.AddAsync(newEntity);
                    }
                }

                auditPeriod.Version++; // Increment version
                _dbContext.tblMdAuditPeriodListTables.Update(auditPeriod);
                var h = new TblMdAuditTemplateHistory()
                {
                    Id = Guid.NewGuid().ToString(),
                    ListAuditCode = auditPeriod.AuditPeriodCode,
                    Action = "Cập nhật dữ liệu",
                    TextContent = null,
                    AuditPeriodListTablesCode = auditPeriod.Code
                };
                await this._dbContext.tblMdAuditTemplateHistory.AddAsync(h);
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
        public async Task<bool> ImportExcelAndUpdateDataDV(IFormFile file, int auditListTablesCode)
        {

            if (file == null || file.Length == 0)
            {
                throw new ArgumentException("File is empty or null.");
            }

            try
            {
                var auditPeriod = await _dbContext.tblMdAuditPeriodListTables
                    .FirstOrDefaultAsync(x => x.Code == auditListTablesCode);

                if (auditPeriod == null)
                {
                    throw new InvalidOperationException($"Audit period with code {auditListTablesCode} not found.");
                }

                var updatedData = new List<AuditTemplateListTablesDataListDto>();
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var stream = new MemoryStream())
                {
                    await file.CopyToAsync(stream);
                    using (var package = new ExcelPackage(stream))
                    {
                        var worksheet = package.Workbook.Worksheets[0];
                        int rowCount = worksheet.Dimension.Rows;

                        for (int row = 9; row <= rowCount; row++)
                        {
                            var orgCodeData = worksheet.Cells[row, 1].Value?.ToString();
                            var listTablesId = worksheet.Cells[row, 3].Value?.ToString();
                            var listTablesName = worksheet.Cells[row, 4].Value?.ToString();

                            if (string.IsNullOrWhiteSpace(orgCodeData) || string.IsNullOrWhiteSpace(listTablesId) || string.IsNullOrWhiteSpace(listTablesName))
                            {
                                continue; // Skip rows with missing essential data
                            }
                            var templateList = await _dbContext.tblMdTemplateListTables
                               .FirstOrDefaultAsync(tem => tem.Code == auditPeriod.ListTablesCode);

                            var temCode = templateList?.Code;
                            var listTable = await _dbContext.TblBuListTables
                                            .FirstOrDefaultAsync(lt => lt.Id == listTablesId &&
                                            lt.Name == listTablesName &&
                                            lt.MgCode == auditPeriod.ListTablesCode);
                            var templateData = await _dbContext.tblMdTemplateListTablesData
                                    .FirstOrDefaultAsync(td => td.TemplateCode == temCode && td.ListTablesCode == listTable.Code &&
                                                    td.OrgCode == orgCodeData);
                            if (listTable == null)
                            {
                                throw new InvalidOperationException($"List table not found for Id: {listTablesId}, Name: {listTablesName}");
                            }

                            var dto = new AuditTemplateListTablesDataListDto
                            {
                                TemplateDataCode = templateData.Code,
                                AuditListTablesCode = auditListTablesCode,
                                OrgCode = orgCodeData,
                                ListTablesCode = listTable.Code,
                                Unit = worksheet.Cells[row, 5].Value?.ToString(),  
                                ExplanationValue = decimal.TryParse(worksheet.Cells[row, 8].Value?.ToString(), out var explanationValue) ? explanationValue : null,
                                ExplanationNote = worksheet.Cells[row, 9].Value?.ToString()
                            };

                            updatedData.Add(dto);
                        }
                    }
                }

                // Update or insert data
                foreach (var dto in updatedData)
                {
                    var existingEntity = await _dbContext.tblMdAuditTemplateListTablesData
                        .FirstOrDefaultAsync(td => td.AuditListTablesCode == dto.AuditListTablesCode &&
                                                   td.TemplateDataCode == dto.TemplateDataCode);

                    if (existingEntity != null)
                    {
                        // Update existing entity
                        _mapper.Map(dto, existingEntity);
                    }
                    else
                    {
                        // Insert new entity
                        var newEntity = _mapper.Map<TblMdAuditTemplateListTablesData>(dto);
                        await _dbContext.tblMdAuditTemplateListTablesData.AddAsync(newEntity);
                    }
                }

                auditPeriod.Version++; // Increment version
                _dbContext.tblMdAuditPeriodListTables.Update(auditPeriod);
                var h = new TblMdAuditTemplateHistory()
                {
                    Id = Guid.NewGuid().ToString(),
                    ListAuditCode = auditPeriod.AuditPeriodCode,
                    Action = "Cập nhật dữ liệu",
                    TextContent = null,
                    AuditPeriodListTablesCode = auditPeriod.Code
                };
                await this._dbContext.tblMdAuditTemplateHistory.AddAsync(h);
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
        public async Task<byte[]> ExportAuditTemplateListTablesData(int auditListTablesCode, string? orgCode)
        {
            try
            {
                var auditPeriod = await _dbContext.tblMdAuditPeriodListTables
                    .Include(x => x.TblMdTemplateListTables)
                    .FirstOrDefaultAsync(x => x.Code == auditListTablesCode);

                if (auditPeriod == null)
                {
                    throw new Exception("Audit period not found");
                }
                var template = await _dbContext.tblMdTemplateListTables
                .FirstOrDefaultAsync(x => x.Code == auditPeriod.ListTablesCode);

                var query = _dbContext.tblMdAuditTemplateListTablesData
                    .Where(x => x.AuditListTablesCode == auditListTablesCode)
                    .Include(x => x.templateListTablesData)
                        .ThenInclude(x => x.Organize)
                    .Include(x => x.templateListTablesData)
                        .ThenInclude(x => x.ListTables)
                    .AsQueryable();
                if (!string.IsNullOrWhiteSpace(orgCode))
                {
                    query = query.Where(x => x.templateListTablesData.OrgCode == orgCode);
                }

                var data = await query.Select(x => new AuditTemplateListTablesDataListDto
                {
                    OrgCode = x.templateListTablesData.OrgCode,
                    OrgName = x.templateListTablesData.Organize.Name,
                    ListTablesId = x.templateListTablesData.ListTables.Id,
                    ListTablesName = x.templateListTablesData.ListTables.Name,
                    Unit = x.Unit ?? "VND",
                    AuditValue = x.AuditValue,
                    AuditExplanation = x.AuditExplanation,
                    ExplanationValue = x.ExplanationValue,
                    ExplanationNote = x.ExplanationNote,
                    OrderNumber = x.templateListTablesData.ListTables.OrderNumber
                }).ToListAsync();

                var sortedData = data.GroupBy(x => x.OrgCode)
                    .SelectMany(group => group.OrderBy(x => x.OrderNumber))
                    .ToList();

                decimal totalAuditValue = sortedData.Sum(x => x.AuditValue ?? 0);
                decimal totalExplanationValue = sortedData.Sum(x => x.ExplanationValue ?? 0);

                // Add total row
                sortedData.Insert(0, new AuditTemplateListTablesDataListDto
                {
                    OrgCode = "Tổng cộng",
                    OrgName = "",
                    ListTablesId = "",
                    ListTablesName = "",
                    Unit = "",
                    AuditValue = totalAuditValue,
                    AuditExplanation = "",
                    ExplanationValue = totalExplanationValue,
                    ExplanationNote = ""
                });

                var additionalInfo = new Dictionary<string, string>
                {
                    { "Người tạo", template.CreateBy ?? "Admin" },
                    { "Người duyệt",  "" }
                };

                return await ExportExtension.ExportToExcel(sortedData, "Sheet 1", $"{auditPeriod.TblMdTemplateListTables.Name}", additionalInfo);
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


