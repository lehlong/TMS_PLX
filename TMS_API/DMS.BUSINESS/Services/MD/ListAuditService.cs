using AutoMapper;
using Common;
using DMS.BUSINESS.Common;
using DMS.BUSINESS.Common.Enum;
using DMS.BUSINESS.Dtos.AD;
using DMS.BUSINESS.Dtos.BU;
using DMS.BUSINESS.Dtos.MD;
using DMS.BUSINESS.Services.BU;
using DMS.COMMON.Common.Enum;
using DMS.CORE;
using DMS.CORE.Entities.BU;
using DMS.CORE.Entities.MD;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Http;
using Microsoft.EntityFrameworkCore;
using NLog.LayoutRenderers.Wrappers;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace DMS.BUSINESS.Services.MD
{
    public interface IListAuditService : IGenericService<TblMdListAudit, ListAuditDto>
    {
        Task<IList<ListAuditDto>> GetAll(BaseMdFilter filter);
        Task<ListAuditDto> GetByCodeAsync(string code);
        Task<OpinionListDto> BuildDataForTreeUnfinished(string timeYear, string auditPeriod);
        Task<ListAuditDto> GetListAuditHistory(string auditCode);
        Task<List<tblMdListAuditHistory>> UpdateListAudit(ListAuditDto audit);
        Task<OrganizeDto> GetOrganizeListAsync(string auditCode);

        Task<List<OpinionStatistic>> getOpinionStatis(string auditCode);
    }
    public class ListAuditService(AppDbContext dbContext, IMapper mapper) : GenericService<TblMdListAudit, ListAuditDto>(dbContext, mapper), IListAuditService
    {
        private readonly IModuleAttachmentService _moduleAttachmentService;
        public override async Task<PagedResponseDto> Search(BaseFilter filter)
        {
            try
            {
                var query = _dbContext.tblMdListAudit.AsQueryable();
                if (!string.IsNullOrWhiteSpace(filter.KeyWord))
                {
                    query = query.Where(x =>
                    x.Name.Contains(filter.KeyWord));
                }
                if (filter.IsActive.HasValue)
                {
                    query = query.Where(x => x.IsActive == filter.IsActive);
                }
                var data = await Paging(query, filter);
                var result = ((IEnumerable)data.Data).Cast<ListAuditDto>().ToList();

                foreach (var i in result)
                {
                    if (i.Status != null || i.Status != "" || string.IsNullOrEmpty(i.Status))
                    {
                        i.Status = GetNameStatus(i.Status);
                    }
                }
                return data;
            }
            catch (Exception ex)
            {
                Status = false;
                Exception = ex;
                return null;
            }
        }

        public async Task<IList<ListAuditDto>> GetAll(BaseMdFilter filter)
        {
            try
            {
                var query = _dbContext.tblMdListAudit.AsQueryable();
                if (filter.IsActive.HasValue)
                {
                    query = query.Where(x => x.IsActive == filter.IsActive);
                }
                var entities = await query.ToListAsync();

                var data = await base.GetAllMd(query, filter);

                foreach (var i in data)
                {
                    if (i.Status != null || i.Status != "" || string.IsNullOrEmpty(i.Status))
                    {
                        i.Status = GetNameStatus(i.Status);
                    }
                }
                return data;

            }
            catch (Exception ex)
            {
                Status = false;
                Exception = ex;
                return null;
            }
        }

        public static string GetNameStatus(string value)
        {
            return value == "01" ? "Khởi tạo" : value == "02" ? "Chờ xác nhận" : value == "03" ? "Đã phê duyệt" : value == "04" ? "Từ chối" : value == "05" ? "Đã hoàn thành" : "Chưa hoàn thành";
        }
        public async Task<ListAuditDto> GetByCodeAsync(string code)
        {
            try
            {
                // Truy vấn tìm bản ghi có mã Code tương ứng
                var entity = await _dbContext.tblMdListAudit
                                .Where(x => x.Code == code)
                                .FirstOrDefaultAsync();

                // Nếu không tìm thấy, trả về null
                if (entity == null)
                {
                    return null;
                }

                // Sử dụng AutoMapper để map từ entity sang DTO
                var dto = _mapper.Map<ListAuditDto>(entity);
                if (!string.IsNullOrEmpty(dto.Status))
                {
                    dto.Status = GetNameStatus(dto.Status);
                }
                return dto;
            }
            catch (Exception ex)
            {
                // Xử lý ngoại lệ nếu có
                Status = false;
                Exception = ex;
                return null;
            }
        }
        public async Task<OpinionListDto> BuildDataForTreeUnfinished(string timeYear, string auditPeriod)
        {
            try
            {
                var lstNode = new List<OpinionListDto>();
                var rootNode = new OpinionListDto()
                {
                    Id = "OPL",
                    PId = "-",
                    Name = "OPL. DANH SÁCH KIẾN NGHỊ CHƯA HOÀN THÀNH",
                    Title = "OPL. DANH SÁCH KIẾN NGHỊ CHƯA HOÀN THÀNH",
                    Key = "OPL"
                };
                lstNode.Add(rootNode);

                var lstAudit = await _dbContext.tblMdListAudit
                    .FirstOrDefaultAsync(a => a.TimeYear == timeYear && a.AuditPeriod == auditPeriod);

                if (lstAudit != null)
                {
                    var opinionDetails = await _dbContext.tblBuOpinionDetail
                        .Where(o => o.MgCode == lstAudit.OpinionCode)
                        .ToListAsync();


                    foreach (var opinionDetail in opinionDetails)
                    {
                        if (opinionDetail.Status == "06")
                        {
                            var opinionUnfinished = await _dbContext.TblBuOpinionLists
                                .Where(u => u.Id == opinionDetail.OpinionCode)
                                .ToListAsync();

                            var opinionOrg = await _dbContext.tblMdOpinionListOrganize
                               .Where(org => opinionUnfinished.Select(x => x.Code).Contains(org.OpinionListCode))
                                .ToListAsync();

                            foreach (var opinion in opinionUnfinished)
                            {
                                var relatedOrgs = opinionOrg
                                    .Where(org => org.OpinionListCode == opinion.Code)
                                    .Select(org => new OpinionListOrganizeDto
                                    {
                                        OrganizeId = org.OrganizeId,
                                    }).ToList();
                                var node = new OpinionListDto()
                                {
                                    Code = opinion.Code,
                                    Id = opinion.Id,
                                    Name = opinion.Name,
                                    PId = rootNode.Id,
                                    OrderNumber = opinion.OrderNumber,
                                    Title = $"{opinion.Id} _ {opinion.Name}",
                                    Key = opinion.Id,
                                    Account = opinion.Account,
                                    MgCode = opinion.MgCode,
                                    OrganizeReferences = relatedOrgs,
                                };
                                lstNode.Add(node);
                            }
                        }
                    }
                }

                // Xây dựng cấu trúc cây
                var nodeDict = lstNode.ToDictionary(n => n.Id);
                foreach (var item in lstNode.Where(n => n.Id != "OPL"))
                {
                    if (nodeDict.TryGetValue(item.PId, out OpinionListDto parentNode))
                    {
                        parentNode.Children ??= new List<OpinionListDto>();
                        parentNode.Children.Add(item);
                    }
                }

                return rootNode;
            }
            catch (Exception ex)
            {
                // Xử lý ngoại lệ
                Status = false;
                Exception = ex;
                return null;
            }
        }

        public async Task<List<tblMdListAuditHistory>> UpdateListAudit(ListAuditDto audit)
        {
            try
            {
                var query = this._dbContext.tblMdListAudit.Find(audit.Code);
                if (query != null)
                {
                    query.Name = audit.Name;
                    query.TimeYear = audit.TimeYear;
                    query.AuditPeriod = audit.AuditPeriod;
                    query.Status = audit.Status;
                    query.ReportNumber = audit.ReportNumber;
                    query.ReportDate = audit.ReportDate;
                    query.StartDate = audit.StartDate;
                    query.EndDate = audit.EndDate;
                    query.OpinionCode = audit.OpinionCode;
                    query.Note = audit.Note;
                    query.FileId = audit.FileId;
                    query.TextContent = audit.TextContent;
                    query.Approver = audit.Approver;
                    var h = new tblMdListAuditHistory()
                    {
                        Id = Guid.NewGuid().ToString(),
                        ListAuditCode = audit.Code,
                        Action = audit.Status == "02" ? "Trình duyệt" : audit.Status == "03" ? "Phê duyệt" : audit.Status == "04" ? "Từ chối" : "Cập nhật thông tin",
                        TextContent = audit.TextContent
                    };
                    this._dbContext.tblMdListAuditHistory.Add(h);
                    await this._dbContext.SaveChangesAsync();
                }
                var lstHistory = this._dbContext.tblMdListAuditHistory.Where(x => x.ListAuditCode == audit.Code).OrderByDescending(x => x.CreateDate).ToList();
                return lstHistory;

            }
            catch (Exception ex)
            {
                this.Status = false;
                this.Exception = ex;
                return new List<tblMdListAuditHistory>();
            }
        }

        public async Task<ListAuditDto> GetListAuditHistory(string auditCode)
        {
            try
            {
                var data = new ListAuditDto();
                var query = await this._dbContext.tblMdListAudit.Where(x => x.Code == auditCode).ToListAsync();

                var audit = query.FirstOrDefault();
                var history = new List<tblMdListAuditHistory>();
                data.Code = audit.Code;
                data.Name = audit.Name;
                data.TimeYear = audit.TimeYear;
                data.AuditPeriod = audit.AuditPeriod;
                data.ReportDate = audit.ReportDate;
                data.ReportNumber = audit.ReportNumber;
                data.Status = audit.Status;
                data.StartDate = audit.StartDate;
                data.EndDate = audit.EndDate;
                data.OpinionCode = audit.OpinionCode;
                data.Note = audit.Note;
                data.FileId = audit.FileId;
                data.IsActive = audit.IsActive;
                data.TextContent = audit.TextContent;
                data.CreateBy = audit.CreateBy;
                data.Approver = audit.Approver;
                history = this._dbContext.tblMdListAuditHistory.Where(x => x.ListAuditCode == auditCode).OrderByDescending(x => x.CreateDate).ToList();
                data.History = history;
                if (!string.IsNullOrEmpty(data.Status))
                {
                    data.Status = GetNameStatus(data.Status);
                }
                return data;

            }
            catch (Exception ex)
            {
                Status = false;
                Exception = ex;
                return null;

            }

        }

        public async Task<OrganizeDto> GetOrganizeListAsync(string auditCode)
        {
            try
            {
                var lstNode = new List<OrganizeDto>();
                var rootNode = new OrganizeDto()
                {
                    Id = "ORG",
                    PId = "-",
                    Name = "1.1. SỞ TÀI CHÍNH",
                    Title = "1.1. SỞ TÀI CHÍNH",
                    Key = "ORG"
                };
                lstNode.Add(rootNode);
                var lstOpinion = _dbContext.TblBuOpinionLists.Where(x => x.MgCode == auditCode).Select(x => x.Code).ToList();
                var lstOrg = _dbContext.tblMdOpinionListOrganize.Where(x => lstOpinion.Contains(x.OpinionListCode)).Select(x => x.OrganizeId).ToList();
                var orgDetails = _dbContext.tblAdOrganize
                .Where(o => lstOrg.Contains(o.Id))
                .ToList();
                foreach (var org in orgDetails)
                {

                    var node = new OrganizeDto()
                    {
                        Id = org.Id,
                        Name = org.Name,
                        PId = rootNode.Id,
                        OrderNumber = org.OrderNumber,
                        Title = $"{org.Id} _ {org.Name}",
                        Key = org.Id,
                    };
                    lstNode.Add(node);
                }
                var nodeDict = lstNode.ToDictionary(n => n.Id);
                foreach (var item in lstNode.Where(n => n.Id != "ORG"))
                {
                    if (nodeDict.TryGetValue(item.PId, out OrganizeDto parentNode))
                    {
                        parentNode.Children ??= new List<OrganizeDto>();
                        parentNode.Children.Add(item);
                    }
                }

                return rootNode;
            }
            catch (Exception ex)
            {
                Status = false;
                Exception = ex;
                return null;
            }
        }

        public async Task<List<OpinionStatistic>> getOpinionStatis(string auditCode)
        {
            try
            {
                var OpinionStatistic = new List<OpinionStatistic>();
                var Details = await _dbContext.tblBuOpinionDetail.Where(d => d.MgCode == auditCode).ToListAsync();
                foreach (var detail in Details)
                {
                    var opn = new OpinionStatistic()
                    {
                        AuditCode = auditCode,
                        OrgCode = detail.OrgCode,
                        Opinion = detail.OpinionCode,
                        Status = GetNameStatus(detail.Status),
                    };
                    OpinionStatistic.Add(opn);
                }

                return OpinionStatistic;
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
