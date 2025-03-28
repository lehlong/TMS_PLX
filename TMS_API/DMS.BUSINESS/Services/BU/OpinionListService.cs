using AutoMapper;
using Common;
using DMS.BUSINESS.Common;
using DMS.BUSINESS.Dtos.AD;
using DMS.BUSINESS.Dtos.BU;
using DMS.BUSINESS.Dtos.MD;
using DMS.BUSINESS.Services.AD;
using DMS.CORE;
using DMS.CORE.Entities.AD;
using DMS.CORE.Entities.BU;
using DMS.CORE.Entities.MD;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc.Filters;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IdentityModel.Tokens.Jwt;
using System.Linq;
using System.Net.WebSockets;
using System.Reflection.Emit;
using System.Security.Claims;
using System.Text;
using System.Threading.Tasks;

namespace DMS.BUSINESS.Services.BU
{
    public interface IOpinionListService : IGenericService<tblBuOpinionList, OpinionListDto>
    {
        Task<OpinionListDto> BuildDataForTree();
        Task<OpinionListDto> BuildDataForTreeWithMgCode(string mgCode);
        Task<OpinionListDto> BuildDataForTreeWithMgCodeAndOrg(string mgCode, string orgCode);
        Task<OpinionListDetailDto> GetOpinionListWithTreeOrganize(object code);
        Task<OpinionDetailDto> GetOpinionDetail(string MgCode,string OrgCode,string OpinionCode);
        Task<OpinionListDto> BuildDataForTreeWithTimeYearAndAuditPeriod(string timeYear, string auditPeriod);
        Task UpdateOrderTree(OpinionListDto moduleDto);
        Task<List<tblBuOpinionDetailHistory>> UpdateOpinionDetail(OpinionDetailDto opd);
        new Task<OpinionListDto> Delete(object code);
        Task<IList<OpinionListDto>> GetAll(BaseFilter filter);
        Task<bool> ImportExcel(IFormFile file, string mgCode);
        //  Task<OpinionListDto> CreateOrderTree(OpinionListDto dto);
    }
    public class OpinionListService(AppDbContext dbContext, IMapper mapper) : GenericService<tblBuOpinionList, OpinionListDto>(dbContext, mapper), IOpinionListService
    {
        
        public async Task<OpinionListDto> BuildDataForTree()
        {
            var lstNode = new List<OpinionListDto>();
            var rootNode = new OpinionListDto() { Id = "OPL", PId = "-", Name = "Danh Mục Ý Kiến", Title = "OPL. DANH MỤC Ý KIẾN", Key = "OPL" };
            lstNode.Add(rootNode);
            var lstAllOpinion = (await this.GetAll()).OrderBy(x => x.OrderNumber).ToList();
            foreach (var Opinion in lstAllOpinion)
            {
                var node = new OpinionListDto()
                {
                    Id = Opinion.Id,
                    Name = Opinion.Name,
                    PId = Opinion.PId,
                    OrderNumber = Opinion.OrderNumber,
                    Title = $"{Opinion.Id} _ {Opinion.Name}",
                    Checked = Opinion.Checked,
                    Account = Opinion.Account,
                    Key = Opinion.Id
                };
                lstNode.Add(node);
            }
            var nodeDict = lstNode.ToDictionary(n => n.Id);
            foreach (var item in lstNode)
            {
                if (item.PId == "-" || !nodeDict.TryGetValue(item.PId, out OpinionListDto parentNode))
                {
                    continue;
                }
                if (parentNode.Children == null)
                {
                    parentNode.Children = new List<OpinionListDto>();
                }
                parentNode.Children.Add(item);
            }
            return rootNode;
        }
        //public async Task<OpinionListDto> CreateOrderTree(OpinionListDto dto)
        //{
        //}

        public async Task UpdateOrderTree(OpinionListDto moduleDto)
        {
            try
            {
                var lstModuleDto = new List<OpinionListDto>();
                var lstModuleUpdate = new List<tblBuOpinionList>();

                this.ConvertNestedToList(moduleDto, ref lstModuleDto);
                if (moduleDto.Children == null || moduleDto.Children.Count == 0)
                {
                    return;
                }
                var numberOrder = 1;
                foreach (var item in lstModuleDto)
                {
                    var module = _mapper.Map<tblBuOpinionList>(item);
                    module.OrderNumber = numberOrder++;
                    lstModuleUpdate.Add(module);
                }
                this._dbContext.UpdateRange(lstModuleUpdate);
                await this._dbContext.SaveChangesAsync();
            }
            catch (Exception ex)
            {
                this.Status = false;
                this.Exception = ex;
            }
        }

        public async Task<List<tblBuOpinionDetailHistory>> UpdateOpinionDetail(OpinionDetailDto opd)
        {
            try
            {
                var i = this._dbContext.tblBuOpinionDetail.Find(opd.Id);
                if (i != null)
                {
                    i.Status = opd.Status;
                    i.ContentOrg = opd.ContentOrg;  
                    i.ContentReport = opd.ContentReport;
                    this._dbContext.tblBuOpinionDetail.Update(i);

                    var h = new tblBuOpinionDetailHistory()
                    {
                        Id = Guid.NewGuid().ToString(),
                        OpinionId = opd.Id,
                        Action = opd.Action == "02" ? "Trình duyệt" : opd.Action == "03" ? "Phê duyệt" : opd.Action == "04" ? "Từ chối" : "Cập nhật thông tin",
                        TextContent = opd.TextContent
                    };
                    this._dbContext.tblBuOpinionDetailHistory.Add(h);
                    await this._dbContext.SaveChangesAsync();
                }
                var lstHistory = this._dbContext.tblBuOpinionDetailHistory.Where(x => x.OpinionId == opd.Id).OrderByDescending(x => x.CreateDate).ToList();
                return lstHistory;
            }
            catch (Exception ex)
            {
                this.Status = false;
                this.Exception = ex;
                return new List<tblBuOpinionDetailHistory>();
            }
        }

        public async Task<OpinionDetailDto> GetOpinionDetail(string MgCode, string OrgCode, string OpinionCode)
        {
            try
            {
                var userRequest = this._dbContext.GetUserRequest();
                var user = this._dbContext.TblAdAccount.Find(userRequest);
                var userOrg = this._dbContext.tblAdOrganize.Find(user.OrganizeCode);

                var data = new OpinionDetailDto();
                var isCreate = await this._dbContext.tblBuOpinionDetail.Where(x => x.OpinionCode == OpinionCode && x.MgCode == MgCode && x.OrgCode == OrgCode).ToListAsync();
                if (isCreate.Any())
                {
                    var d = isCreate.FirstOrDefault();
                    var history = new List<tblBuOpinionDetailHistory>();
                    data.Id = d.Id;
                    data.MgCode = d.MgCode;
                    data.OrgCode = d.OrgCode;
                    data.OpinionCode = d.OpinionCode;
                    data.Status = d.Status;
                    data.ContentOrg = d.ContentOrg;
                    data.ContentReport = d.ContentReport;
                    data.FileId = d.FileId;
                    data.OrgInCharge = d.OrgInCharge;
                    data.CreateBy = d.CreateBy;
                    history = this._dbContext.tblBuOpinionDetailHistory.Where(x => x.OpinionId == d.Id).OrderByDescending(x => x.CreateDate).ToList();
                    data.History = history;
                }
                else
                {
                    var c = new tblBuOpinionDetail()
                    {
                        Id = Guid.NewGuid().ToString(),
                        MgCode = MgCode,
                        OrgCode = OrgCode,
                        OpinionCode = OpinionCode,
                        FileId = Guid.NewGuid().ToString(),
                        Status = "01",
                        OrgInCharge = userOrg.Name
                    };
                    this._dbContext.tblBuOpinionDetail.Add(c);
                    await this._dbContext.SaveChangesAsync();
                    data.Id = c.Id;
                    data.MgCode = c.MgCode;
                    data.OrgCode = c.OrgCode;
                    data.OpinionCode = c.OpinionCode;
                    data.Status = c.Status;
                    data.ContentOrg = c.ContentOrg;
                    data.ContentReport = c.ContentReport;
                    data.FileId = c.FileId;
                    data.CreateBy = c.CreateBy;
                    data.OrgInCharge = c.OrgInCharge;
                }
                return data;
            }
            catch (Exception ex)
            {
                this.Status = false;
                this.Exception = ex;
                return new OpinionDetailDto();
            }
        }

        private void ConvertNestedToList(OpinionListDto node, ref List<OpinionListDto> lstNodeFlat)
        {
            if (node.Id != "OPL")
            {
                lstNodeFlat.Add(node);
            }
            if (node.Children != null && node.Children.Count > 0)
            {
                foreach (var item in node.Children)
                {
                    ConvertNestedToList(item, ref lstNodeFlat);
                }
            }
        }
        public override Task<OpinionListDto> Add(IDto dto)
        {
            var model = dto as OpinionListDto;
            if (string.IsNullOrWhiteSpace(model.PId))
            {
                model.PId = "OPL";
            }
            if (model.OrganizeReferences != null && model.OrganizeReferences.Any())
            {
                foreach (var organize in model.OrganizeReferences)
                {
                    organize.OpinionListCode = model.Code.ToString();
                }
            }
            return base.Add(dto);
        }

        public async new Task<OpinionListDto> Delete(object code)
        {
            try
            {
                var codeGuid = Guid.Parse(code.ToString());
                var recordToDelete = await _dbContext.Set<tblBuOpinionList>().FirstOrDefaultAsync(x => x.Code == codeGuid);
                if (recordToDelete == null)
                {
                    return null;
                }
                var query = _dbContext.Set<tblBuOpinionList>().AsQueryable();
                query = query.Where(x => x.PId == recordToDelete.Id && x.MgCode == recordToDelete.MgCode);
                var recordWithSamePid = await query.ToListAsync();
                if (recordWithSamePid.Count == 0)
                {

                    if (recordToDelete != null)
                    {
                        _dbContext.Remove(recordToDelete);
                        await _dbContext.SaveChangesAsync();
                    }
                    return _mapper.Map<OpinionListDto>(recordToDelete);
                }
                return null;
            }
            catch (Exception ex)
            {
                this.Status = false;
                this.Exception = ex;
                return null;
            }
        }
        public override async Task Update(IDto dto)
        {
            if (dto is OpinionListUpdateDto model)
            {
                var currentObj = await _dbContext.TblBuOpinionLists.Include(x => x.OrganizeReferences).FirstOrDefaultAsync(x => model.Code == x.Code);
                await base.UpdateWithListInside(dto, currentObj);
            }
        }
        public async Task<IList<OpinionListDto>> GetAll(BaseFilter filter)
        {
            try
            {
                var query = this._dbContext.TblBuOpinionLists.AsQueryable();
                if (filter.IsActive.HasValue)
                {
                    query = query.Where(x => x.IsActive == filter.IsActive);
                }
                query = query.OrderByDescending(x => x.CreateDate);
                return _mapper.Map<IList<OpinionListDto>>(await query.ToListAsync());
            }
            catch (Exception ex)
            {
                this.Status = false;
                this.Exception = ex;
                return null;
            }
        }

        public async Task<OpinionListDto> BuildDataForTreeWithMgCode(string mgCode)
        {
            var lstNode = new List<OpinionListDto>();
            var rootNode = new OpinionListDto() { Id = "OPL", PId = "-", Name = "DANH MỤC KIẾN NGHỊ", Title = "OPL. DANH MỤC Ý KIẾN", Key = "OPL" };
            lstNode.Add(rootNode);
            var lstAllOpinion = (await this.GetAll()).Where(x => x.MgCode == mgCode).OrderBy(x => x.OrderNumber).ToList();
            foreach (var Opinion in lstAllOpinion)
            {
                var node = new OpinionListDto()
                {
                    Code = Opinion.Code,
                    Id = Opinion.Id,
                    Name = Opinion.Name,
                    PId = Opinion.PId,
                    OrderNumber = Opinion.OrderNumber,
                    Title = $"{Opinion.Id} _ {Opinion.Name}",
                    Checked = Opinion.Checked,
                    Account = Opinion.Account,
                    Key = Opinion.Id,
                    MgCode = Opinion.MgCode

                };
                lstNode.Add(node);
            }
            var nodeDict = lstNode.ToDictionary(n => n.Id);
            foreach (var item in lstNode)
            {
                if (item.PId == "-" || !nodeDict.TryGetValue(item.PId, out OpinionListDto parentNode))
                {
                    continue;
                }
                if (parentNode.Children == null)
                {
                    parentNode.Children = new List<OpinionListDto>();
                }
                parentNode.Children.Add(item);
            }
            return rootNode;
        }

        public async Task<OpinionListDto> BuildDataForTreeWithMgCodeAndOrg(string mgCode, string orgCode)
        {
            var lstNode = new List<OpinionListDto>();
            var rootNode = new OpinionListDto() { Id = "OPL", PId = "-", Name = "DANH MỤC KIẾN NGHỊ", Title = "OPL. DANH MỤC Ý KIẾN", Key = "OPL" };
            lstNode.Add(rootNode);
            var lstAllOpinion = (await this.GetAll()).Where(x => x.MgCode == mgCode).OrderBy(x => x.OrderNumber).ToList();

            var lstChecked = this._dbContext.tblMdOpinionListOrganize.Where(x => x.OrganizeId == orgCode).Select(x => x.OpinionListCode).Distinct().ToList();
            foreach (var Opinion in lstAllOpinion)
            {
                var node = new OpinionListDto()
                {
                    Code = Opinion.Code,
                    Id = Opinion.Id,
                    Name = Opinion.Name,
                    PId = Opinion.PId,
                    OrderNumber = Opinion.OrderNumber,
                    Title = $"{Opinion.Id} _ {Opinion.Name}",
                    Checked = lstChecked.Contains(Opinion.Code) ? true : false,
                    Account = Opinion.Account,
                    Key = Opinion.Id,
                    MgCode = Opinion.MgCode,
                };
                lstNode.Add(node);
            }
            var nodeDict = lstNode.ToDictionary(n => n.Id);
            foreach (var item in lstNode)
            {
                if (item.PId == "-" || !nodeDict.TryGetValue(item.PId, out OpinionListDto parentNode))
                {
                    continue;
                }
                if (parentNode.Children == null)
                {
                    parentNode.Children = new List<OpinionListDto>();
                }
                parentNode.Children.Add(item);
            }
            return rootNode;
        }

        public async Task<OpinionListDetailDto> GetOpinionListWithTreeOrganize(object code)
        {
            var data = await _dbContext.TblBuOpinionLists.Include(x => x.OrganizeReferences)
                                                         .FirstOrDefaultAsync(x => x.Code == (Guid)code);
            if (data == null) return null;

            var lstNode = new List<OrganizeDto>();
            var rootNode = new OrganizeDto() { Id = "ORG", PId = "-ORG", Name = "STC", Title = "1.1_  STC", Key = "ORG" };
            lstNode.Add(rootNode);

            var lstAllOrganize = await _dbContext.tblAdOrganize.Where(x => x.Id != "ORG" && x.IsActive == true).OrderBy(x => x.OrderNumber).ToListAsync();
            var lstOrgInOpinionList = data.OrganizeReferences.ToDictionary(x => x.OrganizeId, x => x.IsPending);



            foreach (var org in lstAllOrganize)
            {
                var node = new OrganizeDto()
                {
                    Id = org.Id,
                    Name = org.Name,
                    PId = org.PId,
                    OrderNumber = org.OrderNumber,
                    Title = $"{org.Id}_{org.Name}",
                    Key = org.Id,
                    IsPending = false,
                    IsActive = org.IsActive
                };

                if (lstOrgInOpinionList.TryGetValue(org.Id, out bool isPending))
                {
                    node.IsChecked = true;
                    node.IsPending = isPending;
                }

                lstNode.Add(node);
            }

            var nodeDict = lstNode.ToDictionary(n => n.Id);
            foreach (var item in lstNode)
            {
                if (item.PId == "-ORG" || !nodeDict.TryGetValue(item.PId, out OrganizeDto parentNode))
                {
                    continue;
                }

                parentNode.Children ??= [];
                parentNode.Children.Add(item);
            }

            var result = _mapper.Map<OpinionListDetailDto>(data);
            result.TreeOrganize = rootNode;
            return result;
        }

        public async Task<OpinionListDto> BuildDataForTreeWithTimeYearAndAuditPeriod(string timeYear, string auditPeriod)
        {
            var lstNode = new List<OpinionListDto>();
            var rootNode = new OpinionListDto() { Id = "OPL", PId = "-", Name = "DANH SÁCH KIẾN NGHỊ", Title = "OPL. DANH SÁCH KIẾN NGHỊ ", Key = "OPL" };
            lstNode.Add(rootNode);
            var mgOpinionList = await _dbContext.tblMdMgOpinionList
           .FirstOrDefaultAsync(m => m.TimeYear == timeYear && m.AuditPeriod == auditPeriod);

            if (mgOpinionList == null)
            {
                return rootNode; // Return empty tree if no matching MgOpinionList found
            }
            var opinionListWithPendingOrganize = await _dbContext.TblBuOpinionLists
                                                .Where(o => o.MgCode == mgOpinionList.Code)
                                                .OrderBy(o => o.OrderNumber)
                                                .Include(o => o.OrganizeReferences)
                                                .ToListAsync();
            foreach (var opinion in opinionListWithPendingOrganize)
            {
                var node = new OpinionListDto()
                {
                    Code = opinion.Code,
                    Id = opinion.Id,
                    Name = opinion.Name,
                    PId = opinion.PId,
                    OrderNumber = opinion.OrderNumber,
                    Title = $"{opinion.Id} _ {opinion.Name}",
                    Key = opinion.Id,
                    Account = opinion.Account,
                    MgCode = opinion.MgCode,
                    // Thêm thông tin về các TreeOrganize có isPending = true
                    //PendingOrganizes = opinion.OrganizeReferences
                    //    .Where(or => or.IsPending)
                    //    .Select(or => new OrganizeDto
                    //    {
                    //        Id = or.OrganizeId,
                    //        IsPending = or.IsPending,

                    //    })
                    //    .ToList()
                };
                lstNode.Add(node);
            }
            var nodeDict = lstNode.ToDictionary(n => n.Id);
            foreach (var item in lstNode)
            {
                if (item.PId == "-" || !nodeDict.TryGetValue(item.PId, out OpinionListDto parentNode))
                {
                    continue;
                }
                if (parentNode.Children == null)
                {
                    parentNode.Children = new List<OpinionListDto>();
                }
                parentNode.Children.Add(item);
            }
            return rootNode;
        }

        public async Task<bool> ImportExcel(IFormFile file, string mgCode)
        {
            if (file == null || file.Length == 0)
            {
                throw new ArgumentException("File is empty or null.");
            }
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var stream = new MemoryStream())
                {
                    await file.CopyToAsync(stream);
                    using (var package = new ExcelPackage(stream))
                    {
                        var worksheet = package.Workbook.Worksheets[0]; // Lấy sheet đầu tiên

                        int rowCount = worksheet.Dimension.Rows;
                        int orderNumber = 1;

                        for (int row = 3; row <= rowCount; row++) // Bắt đầu từ hàng 3 (bỏ qua tiêu đề)
                        {
                            var id = worksheet.Cells[row, 1].Value?.ToString().Trim();
                            var parentCode = worksheet.Cells[row, 2].Value?.ToString().Trim();
                            var name = worksheet.Cells[row, 3].Value?.ToString().Trim() ?? "";

                            if (string.IsNullOrEmpty(id))
                                continue; // Bỏ qua hàng trống

                            var opinionList = new tblBuOpinionList
                            {
                                Code = Guid.NewGuid(),
                                Id = id,
                                PId = string.IsNullOrEmpty(parentCode) ? "OPL" : parentCode,
                                Name = name,
                                OrderNumber = orderNumber++,
                                MgCode = mgCode
                            };

                            await _dbContext.TblBuOpinionLists.AddAsync(opinionList);
                        }

                        await _dbContext.SaveChangesAsync();
                    }
                }
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

