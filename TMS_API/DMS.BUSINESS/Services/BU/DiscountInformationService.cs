using AutoMapper;
using DMS.BUSINESS.Common;
using DMS.BUSINESS.Dtos.MD;
using DMS.BUSINESS.Models;
using DMS.CORE;
using DMS.CORE.Entities.BU;
using DMS.CORE.Entities.MD;
using Microsoft.AspNetCore.Http;
using Microsoft.EntityFrameworkCore;

using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.ComponentModel;
using System.Drawing;
using System.Reflection;

using SMO;
using OfficeOpenXml.Style;
using System.IO.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NPOI.SS.Util;
using DMS.CORE.Entities.IN;
using System.Reflection.Metadata.Ecma335;
using DMS.BUSINESS.Extentions;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Routing.Template;
using Microsoft.EntityFrameworkCore.Metadata.Internal;

namespace DMS.BUSINESS.Services.BU
{
    public interface IDiscountInformationService : IGenericService<TblMdGoods, GoodsDto>
    {
        Task<string> SaveFileHistory(MemoryStream outFileStream, string headerId);
        void ExportExcel(ref MemoryStream outFileStream, string path, string headerId);
        Task<string> ExportExcelBaoCaoThuLao(string headerId);
        Task<DiscountInformationModel> getAll(string Code);
        Task UpdateDataInput(CompetitorModel model);
        Task<CompetitorModel> getDataInput(string code);
    }
    public class DiscountInformationService(AppDbContext dbContext, IMapper mapper) : GenericService<TblMdGoods, GoodsDto> (dbContext, mapper), IDiscountInformationService
    {
        public async Task<DiscountInformationModel> getAll(string Code)
        {
            try
            {
                var data = new DiscountInformationModel();
                var lstDIL = await _dbContext.TblBuDiscountInformationList.Where(x => x.Code == Code).FirstOrDefaultAsync();
                if(lstDIL != null)
                {

                data.lstDIL = lstDIL;
                var lstMarket = await _dbContext.TblMdMarket.Where(x => x.IsActive == true).OrderBy(x => x.Code).ToListAsync();
                var lstDiscountCompetitor = await _dbContext.TblInDiscountCompetitor.Where(x => x.HeaderCode == Code).ToListAsync();
                var lstInMarketCompetitor = await _dbContext.TblInMarketCompetitor.Where(x => x.HeaderCode == Code).OrderBy(x => x.Code).ToListAsync();
                
                var lstCalculate =  await _dbContext.TblBuInputPrice.Where(x=> x.HeaderId == Code).ToListAsync();

                
                var lstGoods = await _dbContext.TblMdGoods.Where(x => x.IsActive == true).OrderBy(x => x.CreateDate).ToListAsync();
                var discountCompany = await _dbContext.TblInDiscountCompany.Where(x => x.HeaderCode == Code).ToListAsync();
                data.lstGoods = lstGoods;
                var lstCompetitor = await _dbContext.TblMdCompetitor.OrderBy(x => x.Code).ToListAsync();
                data.lstCompetitor = lstCompetitor;

                var orderMarket = 1;
                //var plxna = 1300;
                var z11 = 1961;

                var row1 = new discout
                {
                    colA = "I",
                    colB = "KHO TRUNG TÂM (FOB)",
                    IsBold = true
                };
                foreach (var c in lstCompetitor)
                {
                    row1.gaps.Add(null);
                    row1.cuocVCs.Add(null);
                }

                foreach (var g in lstGoods)
                {
                    var ck = new CK
                    {
                        GoodsCode = g.Code,
                        plxna = discountCompany.FirstOrDefault(d => d.GoodsCode == g.Code).Discount ?? 0,
                    };

                    foreach (var c in lstCompetitor)
                    {
                        var item = lstCalculate.FirstOrDefault(v => v.GoodCode == g.Code);
                        var dt = new DT();  

                        var ck1 = (c.Code == "APP" 
                            ? (0.02m * item.GblV1) + lstDiscountCompetitor.Where(x => x.CompetitorCode == c.Code && x.GoodsCode == g.Code).Sum(x => x.Discount ?? 0 ) 
                            : lstDiscountCompetitor.Where(x => x.CompetitorCode == c.Code && x.GoodsCode == g.Code).Sum(x => x.Discount ?? 0));
                        dt.ckCl.Add(Math.Floor((ck1 / 10)) * 10);
                        dt.ckCl.Add(Math.Round((discountCompany.FirstOrDefault(d => d.GoodsCode == g.Code).Discount ?? 0) - Math.Floor((ck1 / 10)) * 10, 0));
                        dt.code = c.Code;

                        ck.DT.Add(dt);
                    }
                    row1.CK.Add(ck);
                }
                
                data.discount.Add(row1);


                var row2 = new discout
                {
                    colA = "II",
                    colB = "KHO KHÁCH HÀNG (CIF)",
                    IsBold = true
                };
                foreach (var c in lstCompetitor)
                {
                    row2.gaps.Add(null);
                    row2.cuocVCs.Add(null);
                }
                data.discount.Add(row2);


                foreach (var m in lstMarket)
                {
                    var d = new discout
                    {
                        colA = orderMarket.ToString(),
                        colB = m.Name,
                        col1 = m.Gap ?? 0,
                        col4 = m.CuocVCBQ ?? 0,
                    };
                    

                    foreach (var c in lstCompetitor)
                    {
                            //var a = lstInMarketCompetitor.Where(x => x.CompetitorCode == c.Code && x.MarketCode == m.Code).FirstOrDefault();
                        var gap = lstInMarketCompetitor.Where(x => x.CompetitorCode == c.Code && x.MarketCode == m.Code).Sum(x => x.Gap == 0 ? m.Gap + 120 : x.Gap);
                        var cuocVc = c.Code == "APP" ? lstInMarketCompetitor.Where(x => x.CompetitorCode == c.Code && x.MarketCode == m.Code).Sum(x => x.Gap * (decimal)z11 / 1000 ?? 0) : m.CuocVCBQ + 200;
                        d.gaps.Add(Math.Round(gap ?? 0));   
                        d.cuocVCs.Add(cuocVc != null ? Math.Round((decimal)cuocVc, 0) : 0);
                    }
                    foreach (var g in lstGoods)
                    {
                        var ck = new CK
                        {
                            plxna = discountCompany.FirstOrDefault(d => d.GoodsCode == g.Code).Discount - m.CuocVCBQ,
                        };
                        foreach (var c in lstCompetitor)
                        {
                            var item = lstCalculate.FirstOrDefault(v => v.GoodCode == g.Code);
                            var dt = new DT();
                            var discountCompetitor = lstDiscountCompetitor.Where(x => x.CompetitorCode == c.Code && x.GoodsCode == g.Code).Sum(x => x.Discount) ?? 0m;
                            var gap = lstInMarketCompetitor.Where(x => x.CompetitorCode == c.Code && x.MarketCode == m.Code).Sum(x => x.Gap == 0 ? m.Gap + 120 : x.Gap);
                            var cuocVc = c.Code == "APP" ? gap * z11 / 1000 : m.CuocVCBQ + 200;

                            var ck1 = c.Code == "APP"
                            ? 0.02m * item.GblV1 + Math.Round(discountCompetitor , 0) - (cuocVc != null ? Math.Round((decimal)cuocVc, 0) : 0)
                            : discountCompetitor - (cuocVc != null ? Math.Round((decimal)cuocVc, 0) : 0);
                            
                            dt.ckCl.Add(Math.Round((decimal)ck1, 0));
                            dt.ckCl.Add(Math.Round((decimal)((discountCompany.FirstOrDefault(d => d.GoodsCode == g.Code).Discount - m.CuocVCBQ)), 0) - Math.Round((decimal)ck1, 0));
                            dt.code = c.Code;
                            ck.DT.Add(dt);
                        }
                        d.CK.Add(ck);
                    }
                    data.discount.Add(d);
                    orderMarket++;
                }
                return data;
                }
                return data;
            }
            catch (Exception ex)
            {
                this.Status = false;
                this.Exception = ex;
                return new DiscountInformationModel();
            }
        }

        public async Task<CompetitorModel> getDataInput(string code)
        {
            try
            {
                var discoutCompany = await _dbContext.TblInDiscountCompany.Where(x => x.HeaderCode == code).ToListAsync();
                var lstGoods = await _dbContext.TblMdGoods.Where(x => x.IsActive == true).OrderBy(x => x.CreateDate).ToListAsync();
                var lstInMarket = await _dbContext.TblBuInputMarket.OrderBy(x => x.Code).ToListAsync();
                var lstCompetitor = await _dbContext.TblMdCompetitor.OrderBy(x => x.Code).ToListAsync();
                var lstDiscountInformation = await _dbContext.TblInDiscountCompetitor.Where(x => x.HeaderCode == code).ToListAsync();
                var lstInMarketCompetitor = await _dbContext.TblInMarketCompetitor.Where(x => x.HeaderCode == code).ToListAsync();

                List<GOODSs> goodss = new List<GOODSs>();

                foreach (var g in lstGoods)
                {
                    var goods = new GOODSs();
                    goods.Code = g.Code;
                    foreach (var c in lstCompetitor)
                    {
                        goods.HS.Add(new TblInDiscountCompetitor
                        {
                            Code = lstDiscountInformation.Where(x => x.GoodsCode == g.Code && x.CompetitorCode == c.Code).Select(x => x.Code).FirstOrDefault(),
                            HeaderCode = code,
                            GoodsCode = g.Code,
                            Discount = lstDiscountInformation.Where(x => x.GoodsCode == g.Code && x.CompetitorCode == c.Code).Sum(x => x.Discount ?? 0.00M),
                            CompetitorCode = c.Code,
                            IsActive = true,
                        });
                    }
                    goods.DiscountCompany.Add(new TblInDiscountCompany
                    {
                        Code = discoutCompany.FirstOrDefault(x => x.GoodsCode == g.Code).Code,
                        HeaderCode = code,
                        Discount = discoutCompany.FirstOrDefault(x => x.GoodsCode == g.Code).Discount,
                        GoodsCode = discoutCompany.FirstOrDefault(x => x.GoodsCode == g.Code).GoodsCode,
                    });

                    goodss.Add(goods);
                }

                return new CompetitorModel
                {
                    Header = await _dbContext.TblBuDiscountInformationList.Where(x => x.Code == code).FirstOrDefaultAsync(),
                    InMarketCompetitor = lstInMarketCompetitor.Select(x => new TblInMarketCompetitor
                    {
                        Code = x.Code,
                        HeaderCode = code,
                        CompetitorCode = x.CompetitorCode,
                        CompetitorName = x.CompetitorName,
                        MarketCode = x.MarketCode,
                        MarketName = x.MarketName,
                        Gap = x.Gap,

                    }).OrderBy(x => x.CompetitorName).ThenBy(x => x.MarketCode).ToList(),
                    Goodss = goodss,
                };
            }
            catch
            {
                return new CompetitorModel();
            }
        }

        public async Task UpdateDataInput(CompetitorModel model)
        {
            try
            {
                
                _dbContext.TblBuDiscountInformationList.Update(model.Header);
                _dbContext.TblInMarketCompetitor.UpdateRange(model.InMarketCompetitor);

                foreach (var g in model.Goodss)
                {
                    _dbContext.TblInDiscountCompetitor.UpdateRange(g.HS);
                    _dbContext.TblInDiscountCompany.UpdateRange(g.DiscountCompany);
                }

                await _dbContext.SaveChangesAsync();
            }
            catch (DbUpdateConcurrencyException ex)
            {
                Status = false;
                Exception = ex;
            }
        }
     
        public void ExportExcel(ref MemoryStream outFileStream, string path, string headerId)
        {
            try
            {
                FileStream fs = new FileStream(path, FileMode.Open, FileAccess.ReadWrite);
                IWorkbook templateWorkbook = new XSSFWorkbook(fs);
                fs.Close();

                //Define Style
                var styleCellNumber = GetCellStyleNumber(templateWorkbook);

                var font = templateWorkbook.CreateFont();
                font.FontHeightInPoints = 12;
                font.FontName = "Times New Roman";

                ICellStyle styleCellBold = templateWorkbook.CreateCellStyle(); // chữ in đậm
                var fontBold = templateWorkbook.CreateFont();
                fontBold.Boldweight = (short)FontBoldWeight.Bold;
                fontBold.FontHeightInPoints = 12;
                fontBold.FontName = "Times New Roman";

                //Get Data
                var data = getAll(headerId);

                //var startRowPTCK = 0;
                var numbCompetitor = data.Result.lstCompetitor.Count();
                var numbGoods = data.Result.lstGoods.Count();

                var numbCell = (numbCompetitor + 1) * 2 + numbGoods * (numbCompetitor * 2 + 1) + 3;
                //styleCellBold.CloneStyleFrom(sheetPTCK.GetRow(1).Cells[0].CellStyle);
                var cellH = 1;
                ISheet sheetPTCK = templateWorkbook.GetSheetAt(0);
                ExcelNPOIExtention.SetCellValue(sheetPTCK.GetRow(1) ?? sheetPTCK.CreateRow(1), 0, $"TỪ: {data.Result.lstDIL.FDate?.ToString("hh:mm")} ngày {data.Result.lstDIL.FDate?.ToString("dd/MM/yyyy")}", ExcelNPOIExtention.SetCellFreeStyle(templateWorkbook, true, HorizontalAlignment.Center, false, 12));

                // row 1
                #region exp header
                IRow rowCur = ReportUtilities.CreateRow(ref sheetPTCK, 2, numbCell);

                rowCur.Cells[cellH++].SetCellValue("STT");
                rowCur.Cells[cellH++].SetCellValue("Điểm giao hàng");
                rowCur.Cells[cellH].SetCellValue("Cự ly chuyển từ kho trung tâm (Km)");
                sheetPTCK.AddMergedRegion(new CellRangeAddress(2, 3, cellH, cellH+=numbCompetitor));
                cellH++;

                rowCur.Cells[cellH].SetCellValue("Đơn giá cước vận chuyển");
                sheetPTCK.AddMergedRegion(new CellRangeAddress(2, 3, cellH, cellH+=numbCompetitor ));
                cellH++;
                var startR2 = cellH;

                rowCur.Cells[cellH].SetCellValue("CHIẾT KHẤU CÙNG ĐIỂM GIAO");
                sheetPTCK.AddMergedRegion(new CellRangeAddress(2, 2, cellH, cellH+=(numbCompetitor*numbGoods*2 +numbGoods -1)  ));
                cellH++;

                 //row 2
                IRow rowCur2 = ReportUtilities.CreateRow(ref sheetPTCK, 3, numbCell);
                for (var i = 0; i < numbGoods; i++)
                {
                    var lstGoods = data.Result.lstGoods[i];
                    rowCur2.Cells[startR2].SetCellValue(lstGoods.Name);
                    sheetPTCK.AddMergedRegion(new CellRangeAddress(3, 3, startR2, startR2 += (numbCompetitor * 2 )));
                    startR2++;
                }

                //Row 3
                var startR3 = 3;
                IRow rowCur3 = ReportUtilities.CreateRow(ref sheetPTCK, 4, numbCell);
                IRow rowCur4 = ReportUtilities.CreateRow(ref sheetPTCK, 5, numbCell);
                for (var i = 0; i < 2; i++)
                {
                    rowCur3.Cells[startR3].SetCellValue("PLXNA");
                    sheetPTCK.AddMergedRegion(new CellRangeAddress(4, 5, startR3, startR3++));
                    for (var j = 0; j < numbCompetitor; j++)
                    {
                        var lstCompetitor = data.Result.lstCompetitor[j];
                        rowCur3.Cells[startR3].SetCellValue(lstCompetitor.Name);
                        sheetPTCK.AddMergedRegion(new CellRangeAddress(4, 5, startR3, startR3++));
                    }
                }
                var startR4 = 0;
                for (var i = 0; i < numbGoods; i++)
                {
                    rowCur3.Cells[startR3].SetCellValue("PLXNA");
                    sheetPTCK.AddMergedRegion(new CellRangeAddress(4, 5, startR3, startR3++));
                    startR4 = startR3;
                    for (var j = 0; j < numbCompetitor; j++)
                    {
                        var lstCompetitor = data.Result.lstCompetitor[j];
                        rowCur3.Cells[startR3].SetCellValue(lstCompetitor.Name);
                        sheetPTCK.AddMergedRegion(new CellRangeAddress(4, 4, startR3, startR3 += 1));
                        startR3++;

                        rowCur4.Cells[startR4++].SetCellValue("CK");
                        rowCur4.Cells[startR4++].SetCellValue("Chênh lệch so \n với PLX(+/-)");
                    }
                }
                #endregion

                #region exp body
                var startRow = 6;
                var startCell = 1;

                for (var i = 0; i < data.Result.discount.Count(); i++)
                {
                    var dataD = data.Result.discount[i];
                    IRow rowBody = ReportUtilities.CreateRow(ref sheetPTCK, startRow++, numbCell);

                    rowBody.Cells[startCell++].SetCellValue(dataD.colA);
                    rowBody.Cells[startCell++].SetCellValue(dataD.colB);
                    if (i != 1)
                    {
                        rowBody.Cells[startCell].CellStyle = styleCellNumber;
                            rowBody.Cells[startCell++].SetCellValue(Convert.ToDouble(dataD.col1));

                        for (var j = 0; j < dataD.gaps.Count(); j++)
                        {
                            rowBody.Cells[startCell].CellStyle = styleCellNumber;
                            rowBody.Cells[startCell++].SetCellValue((dataD.gaps[j] == 0 || dataD.gaps[j] == null) ? 0 : Convert.ToDouble(dataD.gaps[j]));
                        }
                        rowBody.Cells[startCell].CellStyle = styleCellNumber;
                        rowBody.Cells[startCell++].SetCellValue(dataD.col1 == 0 ? 0 : Convert.ToDouble(dataD.col4));
                        for (var j = 0; j < dataD.cuocVCs.Count(); j++)
                        {
                            rowBody.Cells[startCell].CellStyle = styleCellNumber;
                            rowBody.Cells[startCell++].SetCellValue((dataD.cuocVCs[j] == 0 || dataD.cuocVCs[j] == null) ? 0 : Convert.ToDouble(dataD.cuocVCs[j]));
                        }
                    
                        for (var j = 0; j < data.Result.lstGoods.Count(); j++)
                        {
                            var dataCk = dataD.CK[j];
                            rowBody.Cells[startCell].CellStyle = styleCellNumber;
                            rowBody.Cells[startCell++].SetCellValue((dataD.CK[j].plxna == null || dataD.CK[j].plxna == 0) ? 0 : Convert.ToDouble(dataD.CK[j].plxna));

                            for (var k = 0; k < dataCk.DT.Count(); k++)
                            {
                                for (var z = 0; z < dataCk.DT[k].ckCl.Count(); z++)
                                {
                                    rowBody.Cells[startCell].CellStyle = styleCellNumber;
                                    rowBody.Cells[startCell++].SetCellValue(dataCk.DT[k].ckCl[z] == 0 ? 0 : Convert.ToDouble(dataCk.DT[k].ckCl[z]));
                                }
                            }
                        }
                    }

                    startCell = 1;
                }
                startRow++;
                    ExcelNPOIExtention.SetCellValue(sheetPTCK.GetRow(startRow++) ?? sheetPTCK.CreateRow(startRow++), 26, $"PHÒNG KINH DOANH XĂNG DẦU", ExcelNPOIExtention.SetCellFreeStyle(templateWorkbook, true, HorizontalAlignment.Center, false, 12));

                #endregion
                templateWorkbook.Write(outFileStream);
            }
            catch (Exception ex)
            {
                this.Status = false;
                this.Exception = ex;
            }
        }
        

        public async Task<string> ExportExcelBaoCaoThuLao(string headerId)
        {
            try
            {
                var data = getAll(headerId);

                var templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template", "BaoCaoThuLaoTD.xlsx");

                if (!File.Exists(templatePath))
                {
                    throw new FileNotFoundException("Không tìm thấy template Excel.", templatePath);
                }
                using var file = new FileStream(templatePath, FileMode.Open, FileAccess.Read, FileShare.Read);
                IWorkbook workbook = new XSSFWorkbook(file);

                var styles = new
                {
                    FreeText = ExcelNPOIExtention.SetCellFreeStyle(workbook, true, HorizontalAlignment.Center, true, 16),
                    Text = ExcelNPOIExtention.SetCellStyleText(workbook, false, HorizontalAlignment.Left, true),
                    TextRight = ExcelNPOIExtention.SetCellStyleText(workbook, false, HorizontalAlignment.Right, true),
                    TextBold = ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Left, true),
                    TextCenter = ExcelNPOIExtention.SetCellStyleText(workbook, false, HorizontalAlignment.Center, false),
                    TextCenterBold = ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false),
                    Number = ExcelNPOIExtention.SetCellStyleNumber(workbook, false, HorizontalAlignment.Right, true),
                    NumberBold = ExcelNPOIExtention.SetCellStyleNumber(workbook, true, HorizontalAlignment.Right, true),
                };
                var sheetBaoCao = workbook.GetSheetAt(0);

                #region thông báo   
                int rowIndex = 6;
                var fDate = data.Result.lstDIL.FDate.Value.ToString("dd.MM.yyyy");
                var d = data.Result.discount[0];
                
                var row = sheetBaoCao.GetRow(rowIndex) ?? sheetBaoCao.CreateRow(rowIndex);

                ExcelNPOIExtention.SetCellValueText(row, 0, fDate, styles.Text);
                ExcelNPOIExtention.SetCellValueText(row, 1, "31.12.9999", styles.Text);
                ExcelNPOIExtention.SetCellValueNumber(row, 2, 2810, styles.Number);
                ExcelNPOIExtention.SetCellValueText(row, 3, "Công ty Xăng dầu Nghệ An", styles.Text);
                ExcelNPOIExtention.SetCellValueNumber(row, 4, "900001", styles.Number);
                ExcelNPOIExtention.SetCellValueText(row, 5, "Petrolimex", styles.Text);
                ExcelNPOIExtention.SetCellValueText(row, 6, "Nghệ An", styles.Text);
                var indexCk = 7;
                foreach (var ck in d.CK)
                {
                    if (ck.GoodsCode != "0601005")
                    {
                        ExcelNPOIExtention.SetCellValueNumber(row, indexCk++, ck.plxna, styles.Number);
                        ExcelNPOIExtention.SetCellValueNumber(row, indexCk++, ck.plxna, styles.Number);
                    }
                }
                ExcelNPOIExtention.SetCellValueText(row, indexCk++, "Bến thủy/Nghi Hương", styles.Text);
                ExcelNPOIExtention.SetCellValueText(row, indexCk++, "CK vùng 2", styles.Text);
                rowIndex++;

                foreach (var i in data.Result.lstCompetitor)
                {
                    var khoBenBan = i.Code == "APP" ? "Nghi Sơn" : "Vũng Áng";
                    var MaDauMoi = i.Code == "APP" ? "900003" : "900002";
                    var text = d.IsBold ? styles.TextBold : styles.Text;
                    var number = d.IsBold ? styles.NumberBold : styles.Number;
                    var row1 = sheetBaoCao.GetRow(rowIndex) ?? sheetBaoCao.CreateRow(rowIndex);

                    ExcelNPOIExtention.SetCellValueText(row1, 0, fDate, styles.Text);
                    ExcelNPOIExtention.SetCellValueText(row1, 1, "31.12.9999", styles.Text);
                    ExcelNPOIExtention.SetCellValueNumber(row1, 2, 2810, styles.Number);
                    ExcelNPOIExtention.SetCellValueText(row1, 3, "Công ty Xăng dầu Nghệ An", styles.Text);
                    ExcelNPOIExtention.SetCellValueNumber(row1, 4,  MaDauMoi, styles.Number);
                    ExcelNPOIExtention.SetCellValueText(row1, 5, i.Name, styles.Text);
                    ExcelNPOIExtention.SetCellValueText(row1, 6, "Hà Tinh/Nghệ An/Thanh Hóa", styles.Text);
                    indexCk = 7;
                    foreach (var ck in d.CK)
                    {
                        if (ck.GoodsCode != "0601005")
                        {
                            var indexCkCl = 0;
                            //foreach (var ckDt in ck.DT)
                            //{
                                if(rowIndex == 7)
                                {
                                    ExcelNPOIExtention.SetCellValueNumber(row1, indexCk++, ck.DT[0].ckCl[0], styles.Number);
                                    ExcelNPOIExtention.SetCellValueNumber(row1, indexCk++, ck.DT[0].ckCl[0], styles.Number);
                                }
                                else
                                {
                                    ExcelNPOIExtention.SetCellValueNumber(row1, indexCk++, ck.DT[1].ckCl[0], styles.Number);
                                    ExcelNPOIExtention.SetCellValueNumber(row1, indexCk++, ck.DT[1].ckCl[0], styles.Number);
                                }
                            //}
                        }
                        //var ckDt = ck.DT[0];
                    }
                    ExcelNPOIExtention.SetCellValueText(row1, 13, khoBenBan, styles.Text);
                    ExcelNPOIExtention.SetCellValueText(row1, 14, "CK vùng 2", styles.Text);
                    rowIndex++;
                }

                #endregion

                var folderPath = Path.Combine($"Uploads/Excel/{DateTime.Now.ToString("yyyy/MM/dd")}");
                if (!Directory.Exists(folderPath))
                {
                    Directory.CreateDirectory(folderPath);
                }
                var fileName = $"BCThuLaoTD_{DateTime.Now:ddMMyyyy_HHmmss}.xlsx";
                var outputPath = Path.Combine(Directory.GetCurrentDirectory(), folderPath, fileName);

                using var outFile = new FileStream(outputPath, FileMode.Create, FileAccess.Write);
                workbook.Write(outFile);

                _dbContext.TblBuHistoryDownload.Add(new TblBuHistoryDownload
                {
                    Code = Guid.NewGuid().ToString(),
                    HeaderCode = headerId,
                    Name = fileName,
                    Type = "xlsx",
                    Path = $"{folderPath}/{fileName}",
                });
                await _dbContext.SaveChangesAsync();

                return $"{folderPath}/{fileName}";
                //templateWorkbook.Write(outFileStream);
            }
            catch (Exception ex)
            {
                return null;
            }
        }


        public ICellStyle GetCellStyleNumber(IWorkbook templateWorkbook)
        {
            ICellStyle styleCellNumber = templateWorkbook.CreateCellStyle();
            styleCellNumber.DataFormat = templateWorkbook.CreateDataFormat().GetFormat("#,##0");
            return styleCellNumber;
        }

        public async Task<string> SaveFileHistory(MemoryStream outFileStream, string headerId)
        {
            byte[] data = outFileStream.ToArray();
            var path = "";
            using (MemoryStream memoryStream = new MemoryStream(data))
            {
                IFormFile file = ConvertMemoryStreamToIFormFile(memoryStream, "example.txt");
                var folderName = Path.Combine($"Uploads/{DateTime.Now.Year}/{DateTime.Now.Month}/{DateTime.Now.Day}");
                //var folderName = Path.Combine("D:\\dowloads\\xuatexcel");
                if (!Directory.Exists(folderName))
                {
                    Directory.CreateDirectory(folderName);
                }
                var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
                if (file.Length > 0)
                {
                    var fileName = $"{DateTime.Now.Day}{DateTime.Now.Month}{DateTime.Now.Year}_{DateTime.Now.Hour}{DateTime.Now.Minute}{DateTime.Now.Second}_PhanTichChietKhau.xlsx";
                    var fullPath = Path.Combine(pathToSave, fileName);
                    using (var stream = new FileStream(fullPath, FileMode.Create))
                    {
                        file.CopyTo(stream);
                    }
                    path = $"Uploads/{DateTime.Now.Year}/{DateTime.Now.Month}/{DateTime.Now.Day}/{fileName}";
                    _dbContext.TblBuHistoryDownload.Add(new TblBuHistoryDownload
                    {
                        Code = Guid.NewGuid().ToString(),
                        HeaderCode = headerId,
                        Name = fileName,
                        Type = "xlsx",
                        Path = path
                    });
                    await _dbContext.SaveChangesAsync();
                }
            }
            return path;
        }
       
        public static IFormFile ConvertMemoryStreamToIFormFile(MemoryStream memoryStream, string fileName)
        {
            memoryStream.Position = 0; // Reset the stream position to the beginning
            IFormFile formFile = new FormFile(memoryStream, 0, memoryStream.Length, "file", fileName)
            {
                Headers = new HeaderDictionary(),
                ContentType = "application/octet-stream"
            };
            return formFile;

        }
    
        
    
    }


}
