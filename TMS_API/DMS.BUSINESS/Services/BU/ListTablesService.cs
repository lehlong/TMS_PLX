using AutoMapper;
using Common;
using DMS.BUSINESS.Common;
using DMS.BUSINESS.Dtos.BU;
using DMS.BUSINESS.Dtos.MD;
using DMS.CORE;
using DMS.CORE.Entities.BU;
using DMS.CORE.Entities.MD;
using Microsoft.AspNetCore.Http;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMS.BUSINESS.Services.BU
{
    public interface IListTablesService : IGenericService<TblBuListTables, ListTablesDto>
    {
        Task<ListTablesDto> BuildDataForTree();
        Task<ListTablesDto> BuildDataForTreeWithMgCode(string mgCode);
        Task UpdateOrderTree(ListTablesDto moduleDto);
        new Task<ListTablesDto> Delete(object code);
        Task<IList<ListTablesDto>> GetAll(BaseFilter filter);
        Task<bool> ImportExcel(IFormFile file, string mgCode);
    }
    public class ListTablesService(AppDbContext dbContext, IMapper mapper) : GenericService<TblBuListTables, ListTablesDto>(dbContext, mapper), IListTablesService
    {
        public async Task<ListTablesDto> BuildDataForTree()
        {
            var lstNode = new List<ListTablesDto>();
            var rootNode = new ListTablesDto() { Id = "LTB", PId = "-", Name = "Danh Mục Bảng Biểu", Title = "LTB. DANH MỤC BẢNG BIỂU", Key = "LTB" };
            lstNode.Add(rootNode);
            var lstAllTables = (await this.GetAll()).OrderBy(x => x.OrderNumber).ToList();
            foreach (var tables in lstAllTables)
            {
                var node = new ListTablesDto()
                {
                    Id = tables.Id,
                    Name = tables.Name,
                    PId = tables.PId,
                    OrderNumber = tables.OrderNumber,
                    Title = $"{tables.Id} _ {tables.Name}",
                    IsChecked = tables.IsChecked,
                    Key = tables.Id,
                    CurrencyCode = tables.CurrencyCode
                };
                lstNode.Add(node);
            }
            var nodeDict = lstNode.ToDictionary(n => n.Id);
            foreach (var item in lstNode)
            {
                if (item.PId == "-" || !nodeDict.TryGetValue(item.PId, out ListTablesDto parentNode))
                {
                    continue;
                }
                if (parentNode.Children == null)
                {
                    parentNode.Children = new List<ListTablesDto>();
                }
                parentNode.Children.Add(item);
            }
            return rootNode;
        }
        public async Task<ListTablesDto> BuildDataForTreeWithMgCode(string mgCode)
        {
            var lstNode = new List<ListTablesDto>();
            var rootNode = new ListTablesDto() { Id = "LTB", PId = "-", Name = "Danh Mục Chỉ Tiêu", Title = "LTB. DANH MỤC CHỈ TIÊU", Key = "LTB" };
            lstNode.Add(rootNode);
            var lstAllTables = (await this.GetAll())
                .Where(x => x.MgCode == mgCode)
                .OrderBy(x => x.OrderNumber).ToList();
            foreach (var tables in lstAllTables)
            {
                var node = new ListTablesDto()
                {
                    Code = tables.Code,
                    Id = tables.Id,
                    Name = tables.Name,
                    PId = tables.PId,
                    OrderNumber = tables.OrderNumber,
                    Title = $"{tables.Id} _ {tables.Name}",
                    IsChecked = tables.IsChecked,
                    Key = tables.Id,
                    MgCode = tables.MgCode,
                    CurrencyCode = tables.CurrencyCode,
                    IsLeaf = true
                };
                lstNode.Add(node);
            }
            var nodeDict = lstNode.ToDictionary(n => n.Id);
            foreach (var item in lstNode)
            {
                if (item.PId == "-" || !nodeDict.TryGetValue(item.PId, out ListTablesDto parentNode))
                {
                    continue;
                }
                if (parentNode.Children == null)
                {
                    parentNode.Children = new List<ListTablesDto>();
                }
                parentNode.Children.Add(item);
                parentNode.IsLeaf = false;
            }
            foreach (var node in lstNode)
            {
                if (node.Children == null || node.Children.Count == 0)
                {
                    node.IsLeaf = true;
                }
            }
            return rootNode;
        }
        public async Task UpdateOrderTree(ListTablesDto moduleDto)
        {
            try
            {
                var lstModuleDto = new List<ListTablesDto>();
                var lstModuleUpdate = new List<TblBuListTables>();

                this.ConvertNestedToList(moduleDto, ref lstModuleDto);
                if (moduleDto.Children == null || moduleDto.Children.Count == 0)
                {
                    return;
                }
                var numberOrder = 1;
                foreach (var item in lstModuleDto)
                {
                    var module = _mapper.Map<TblBuListTables>(item);
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

        private void ConvertNestedToList(ListTablesDto node, ref List<ListTablesDto> lstNodeFlat)
        {
            if (node.Id != "LTB")
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
        public override Task<ListTablesDto> Add(IDto dto)
        {
            var model = dto as ListTablesDto;
            if (string.IsNullOrWhiteSpace(model.PId))
            {
                model.PId = "LTB";
            }
            return base.Add(dto);
        }

        public async new Task<ListTablesDto> Delete(object code)
        {
            try
            {
                var codeGuid = Guid.Parse(code.ToString());
                var recordToDelete = await _dbContext.Set<TblBuListTables>().FirstOrDefaultAsync(x => x.Code == codeGuid);
                if (recordToDelete == null)
                {
                    return null;
                }
                var query = _dbContext.Set<TblBuListTables>().AsQueryable();
                query = query.Where(x => x.PId == recordToDelete.Id && x.Code == recordToDelete.Code);
                var recordWithSamePid = await query.ToListAsync();
                if (recordWithSamePid.Count == 0)
                {
                    if (recordToDelete != null)
                    {
                        _dbContext.Remove(recordToDelete);
                        await _dbContext.SaveChangesAsync();
                    }
                    return _mapper.Map<ListTablesDto>(recordToDelete);
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
        public async Task<IList<ListTablesDto>> GetAll(BaseFilter filter)
        {
            try
            {
                var query = this._dbContext.TblBuListTables.AsQueryable();
                if (filter.IsActive.HasValue)
                {
                    query = query.Where(x => x.IsActive == filter.IsActive);
                }
                query = query.OrderByDescending(x => x.CreateDate);
                return _mapper.Map<IList<ListTablesDto>>(await query.ToListAsync());
            }
            catch (Exception ex)
            {
                this.Status = false;
                this.Exception = ex;
                return null;
            }
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
                            var name = worksheet.Cells[row, 3].Value?.ToString().Trim()??"";
                            var currencyCode = worksheet.Cells[row, 4].Value?.ToString().Trim();

                            if (string.IsNullOrEmpty(id))
                                continue; // Bỏ qua hàng trống

                            var listTable = new TblBuListTables
                            {
                                Code = Guid.NewGuid(),
                                Id = id,
                                PId = string.IsNullOrEmpty(parentCode) ? "LTB" : parentCode,
                                Name = name,
                                OrderNumber = orderNumber++,
                                CurrencyCode = currencyCode,
                                MgCode = mgCode
                            };

                            await _dbContext.TblBuListTables.AddAsync(listTable);
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

