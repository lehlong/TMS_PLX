using AutoMapper;
using Common;
using DMS.BUSINESS.Common;
using DMS.BUSINESS.Dtos.MD;
using DMS.CORE;
using DMS.CORE.Entities.MD;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DMS.BUSINESS.Services.MD
{
    public interface ICuocVanChuyenService : IGenericService<TblMdCuocVanChuyen, CuocVanChuyenDto>
    {
        Task<IList<CuocVanChuyenDto>> GetAll(BaseMdFilter filter);
        Task<IList<CuocVanChuyenDto>> ImportExcelData(string filePath, string  headerCode);
        Task<PagedResponseDto> SearchById(BaseFilter filter, string id);
    }

    class CuocVanChuyenService(AppDbContext dbContext, IMapper mapper) : GenericService<TblMdCuocVanChuyen, CuocVanChuyenDto>(dbContext, mapper), ICuocVanChuyenService
    {
        public async Task<IList<CuocVanChuyenDto>> GetAll(BaseMdFilter filter)
        {
            try
            {
                var query = _dbContext.TblMdCuocVanChuyen.AsQueryable();
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
        public async Task<IList<CuocVanChuyenDto>> ImportExcelData(string filePath, string headerCode)
        {
            List<TblMdCuocVanChuyen> data = new List<TblMdCuocVanChuyen>();
            int rowIndex = 1;
            Status = true;

            try
            {
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(filePath, false))
                {
                    WorkbookPart workbookPart = doc.WorkbookPart;
                    Sheet sheet = workbookPart.Workbook.Sheets.GetFirstChild<Sheet>();
                    WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                    SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                    var headerRow = sheetData.Elements<Row>().FirstOrDefault();
                    var headerCells = headerRow.Elements<Cell>().Select(c => GetCellValue(c, workbookPart)).ToList();

                    var rows = sheetData.Elements<Row>().Skip(2);
                    foreach (var row in rows)
                    {
                        var cells = FillMissingCells(row, headerCells.Count > row.Elements<Cell>().Count() ? headerCells.Count : row.Elements<Cell>().Count());
                        bool isRowEmpty = cells.All(c => string.IsNullOrWhiteSpace(GetCellValue(c, workbookPart)));
                        if (isRowEmpty) continue;

                        try
                        {
                            TblMdCuocVanChuyen record = new TblMdCuocVanChuyen
                            {
                                Code = Guid.NewGuid().ToString(),
                                VSART = headerCells.Contains("VSART") && headerCells.IndexOf("VSART") < cells.Count ? GetCellValue(cells[headerCells.IndexOf("VSART")], workbookPart) : null,
                                TDLNR = headerCells.Contains("TDLNR") && headerCells.IndexOf("TDLNR") < cells.Count ? GetCellValue(cells[headerCells.IndexOf("TDLNR")], workbookPart) : null,
                                KNOTA = headerCells.Contains("KNOTA") && headerCells.IndexOf("KNOTA") < cells.Count ? GetCellValue(cells[headerCells.IndexOf("KNOTA")], workbookPart) : null,
                                OIGKNOTD = headerCells.Contains("OIGKNOTD") && headerCells.IndexOf("OIGKNOTD") < cells.Count ? GetCellValue(cells[headerCells.IndexOf("OIGKNOTD")], workbookPart) : null,
                                MATNR = headerCells.Contains("MATNR") && headerCells.IndexOf("MATNR") < cells.Count ? GetCellValue(cells[headerCells.IndexOf("MATNR")], workbookPart) : null,
                                VRKME = headerCells.Contains("VRKME") && headerCells.IndexOf("VRKME") < cells.Count ? GetCellValue(cells[headerCells.IndexOf("VRKME")], workbookPart) : null,
                                KBETR = headerCells.Contains("KBETR") && headerCells.IndexOf("KBETR") < cells.Count && decimal.TryParse(GetCellValue(cells[headerCells.IndexOf("KBETR")], workbookPart), out var amount) ? amount : 0,
                                KONWA = headerCells.Contains("KONWA") && headerCells.IndexOf("KONWA") < cells.Count ? GetCellValue(cells[headerCells.IndexOf("KONWA")], workbookPart) : null,
                                KPEIN = headerCells.Contains("KPEIN") && headerCells.IndexOf("KPEIN") < cells.Count ? GetCellValue(cells[headerCells.IndexOf("KPEIN")], workbookPart) : null,
                                KMEIN = headerCells.Contains("KMEIN") && headerCells.IndexOf("KMEIN") < cells.Count ? GetCellValue(cells[headerCells.IndexOf("KMEIN")], workbookPart) : null,
                                DATAB = headerCells.Contains("DATAB") && headerCells.IndexOf("DATAB") < cells.Count ? GetCellValue(cells[headerCells.IndexOf("DATAB")], workbookPart) : null,
                                DATBI = headerCells.Contains("DATBI") && headerCells.IndexOf("DATBI") < cells.Count ? GetCellValue(cells[headerCells.IndexOf("DATBI")], workbookPart) : null,
                                HeaderCode = headerCode,
                                IsActive = true,
                                RowIndex = rowIndex
                            };
                            data.Add(record);
                            rowIndex++;
                        }
                        catch (Exception ex)
                        {
                            Status = false;
                            Console.WriteLine($"Error processing row {rowIndex}: {ex.Message}");
                            continue;
                        }
                    }
                }

                if (data.Any())
                {
                    await _dbContext.TblMdCuocVanChuyen.AddRangeAsync(data);
                    await _dbContext.SaveChangesAsync();
                }
            }
            catch (Exception ex)
            {
                Status = false;
                Exception = ex;
                return null;
            }

            return data.Select(d => new CuocVanChuyenDto
            {
                Code = d.Code,
                VSART = d.VSART,
                TDLNR = d.TDLNR,
                KNOTA = d.KNOTA,
                OIGKNOTD = d.OIGKNOTD,
                MATNR = d.MATNR,
                VRKME = d.VRKME,
                KBETR = d.KBETR,
                KONWA = d.KONWA,
                KPEIN = d.KPEIN,
                KMEIN = d.KMEIN,
                DATAB = d.DATAB,
                DATBI = d.DATBI,
                IsActive = d.IsActive,
                RowIndex = d.RowIndex
            }).ToList();
        }

        public async Task<PagedResponseDto> SearchById(BaseFilter filter, string id)
        {
            try
            {
                var query = _dbContext.TblMdCuocVanChuyen.AsQueryable().Where(x => x.HeaderCode == id); ;

                if (!string.IsNullOrWhiteSpace(filter.KeyWord))
                {
                    query = query.Where(x =>
                    x.VSART.Contains(filter.KeyWord));
                }
                if (filter.IsActive.HasValue)
                {
                    query = query.Where(x => x.IsActive == filter.IsActive);
                }
                query = query.OrderBy(x => x.RowIndex);
                return await Paging(query, filter);
            }
            catch (Exception ex)
            {
                Status = false;
                Exception = ex;
                return null;
            }
        }

        private string GetCellValue(Cell cell, WorkbookPart workbookPart)
        {
            if (cell == null)
                return null;

            string value = cell.CellValue?.InnerText;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                var sharedStringTable = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                if (sharedStringTable != null && int.TryParse(value, out int index))
                {
                    value = sharedStringTable.SharedStringTable.ElementAt(index).InnerText;
                }
            }

            return string.IsNullOrWhiteSpace(value) ? null : value.Trim();
        }
        private List<Cell?> FillMissingCells(Row row, int totalCells)
        {
            List<Cell?> filledCells = new List<Cell?>();
            int cellIndex = 0;

            var cells = row.Elements<Cell>().ToList();

            foreach (var cell in cells)
            {
                int columnIndex = GetColumnIndex(cell.CellReference);
                while (cellIndex < columnIndex)
                {
                    filledCells.Add(null);
                    cellIndex++;
                }
                filledCells.Add(cell);
                cellIndex++;
            }

            while (filledCells.Count < totalCells)
            {
                filledCells.Add(null);
            }

            return filledCells;
        }

        private int GetColumnIndex(string cellReference)
        {
            string columnReference = new string(cellReference.Where(c => Char.IsLetter(c)).ToArray());
            int columnIndex = 0;
            int factor = 1;

            for (int i = columnReference.Length - 1; i >= 0; i--)
            {
                columnIndex += (columnReference[i] - 'A' + 1) * factor;
                factor *= 26;
            }
            return columnIndex - 1;
        }
    }
}
