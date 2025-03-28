using AutoMapper;
using Common;
using DMS.BUSINESS.Common;
using DMS.BUSINESS.Dtos.AD;
using DMS.BUSINESS.Dtos.BU;
using DMS.BUSINESS.Dtos.MD;
using DMS.BUSINESS.Services.BU;
using DMS.BUSINESS.Services.BU.Attachment;
using DMS.CORE;
using DMS.CORE.Entities.BU;
using DMS.CORE.Entities.MD;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.EntityFrameworkCore;
using PROJECT.Service.Extention;
using System.Linq;
using System.Reflection.Emit;

using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using System.Data;
using System.Net.Http.Headers;
using Microsoft.AspNetCore.Http.Extensions;
using Microsoft.AspNetCore.Http;
using System.Net.Http;
using DocumentFormat.OpenXml.Drawing.Diagrams;


namespace DMS.BUSINESS.Services.Report
{
    public interface IReportService : IGenericService<TblMdTemplateReportMapping, TemplateReportMappingDto>
    {
        //Task<object> Upload(IFormFile file, string? moduleType, Guid? referenceId, string? auditValue, string? yearValue);
        Task ProcessFile(string path, string audit, string year);
        Task SaveReportTemplate(string fullPath, string auditValue, string yearValue, string fileOldName, string fileExt, string fileSize, string fileName, string netWorkPath);
        Task<List<TblMdTemplateReport>> GetListTemplate(string year, string audit);
        Task<List<TblMdTemplateReportElement>> GetListElement(string fileId);
        Task SaveTemplateReport(TemplateReportMappingDto dto);
        Task<OrganizeDto> GetListOrg(string fileId, string textElement);
        Task<OpinionListDto> GetListOpinion(string fileId, string textElement, string OrgCode, string timeYear, string auditPeriod);
        Task<TblMdTemplateReport> GetTemplate(string id, string year, string audit);

    }
    public class ReportService(AppDbContext dbContext, IMapper mapper, IOpinionListService _serviceOpinion) : GenericService<TblMdTemplateReportMapping, TemplateReportMappingDto>(dbContext, mapper), IReportService
    {

        private readonly IOpinionListService _serviceOpinion = _serviceOpinion;

        public async Task<List<TblMdTemplateReport>> GetListTemplate(string year, string audit)
        {
            try
            {
                var result = _dbContext.tblMdTemplateReport.ToList();
                if (!string.IsNullOrEmpty(year)) result = result.Where(x => x.Year == year).ToList();
                if (!string.IsNullOrEmpty(audit)) result = result.Where(x => x.AuditPeriod == audit).ToList();
                return result;
            }
            catch (Exception ex)
            {
                return new List<TblMdTemplateReport>();
            }

        }
        public async Task<List<TblMdTemplateReportElement>> GetListElement(string fileId)
        {
            try
            {
                return _dbContext.tblMdTemplateReportElement.Where(x => x.FileId == fileId).ToList();
            }
            catch (Exception ex)
            {
                return new List<TblMdTemplateReportElement>();
            }
        }

        public async Task ProcessFile(string path, string audit, string year)
        {
            List<string> lstTextElement = new List<string>();
            WordDocumentService wordDocumentService = new WordDocumentService();
            using (WordprocessingDocument doc = WordprocessingDocument.Open(path, true))
            {
                lstTextElement = wordDocumentService.FindTextElement(doc);
            }

            var headerData = _dbContext.tblMdListAudit.Where(x => x.TimeYear == year && x.AuditPeriod == audit).FirstOrDefault();
            var dataDetail = _dbContext.tblBuOpinionDetail.Where(x => headerData.Code.Contains(x.MgCode)).ToList();

            using (WordprocessingDocument doc = WordprocessingDocument.Open(path, true))
            {
                MainDocumentPart mainPart = doc.MainDocumentPart;
                Body body = mainPart.Document.Body;


                foreach (var text in lstTextElement)
                {
                    var textReplace = text.Replace("##", "").Replace("@@", "").Split("_").ToList();

                    #region Gen kiến nghị
                    if (textReplace.Count() == 3)
                    {
                        var dataTreeOpinion = _dbContext.TblBuOpinionLists.Where(x => x.MgCode == headerData.OpinionCode).ToList();
                        var nameOpinion = dataTreeOpinion.Where(x => x.Id == textReplace[1]).FirstOrDefault().Name;
                        var dataFill = dataDetail.Where(x => x.OrgCode == textReplace[0] && x.OpinionCode == textReplace[1]).FirstOrDefault();
                        if (textReplace[2] == "N")
                        {
                            wordDocumentService.ReplaceStringInWordDocumennt(doc, text, nameOpinion);
                        }
                        else
                        {
                            wordDocumentService.ReplaceStringInWordDocumennt(doc, text, dataFill.ContentReport);
                        }
                    }
                    #endregion

                    #region Gen bảng biểu
                    if (textReplace.Count() == 2)
                    {
                        Paragraph paragraph = body.Descendants<Paragraph>()
                                          .FirstOrDefault(p => p.InnerText.Contains(text));

                        if (paragraph != null)
                        {
                            Table table = new Table();
                            TableProperties tblProperties = new TableProperties(
                                new TableBorders(
                                    new TopBorder { Val = new EnumValue<BorderValues>(BorderValues.BasicThinLines), Size = 12 },
                                    new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.BasicThinLines), Size = 12 },
                                    new LeftBorder { Val = new EnumValue<BorderValues>(BorderValues.BasicThinLines), Size = 12 },
                                    new RightBorder { Val = new EnumValue<BorderValues>(BorderValues.BasicThinLines), Size = 12 },
                                    new InsideHorizontalBorder { Val = new EnumValue<BorderValues>(BorderValues.BasicThinLines), Size = 12 },
                                    new InsideVerticalBorder { Val = new EnumValue<BorderValues>(BorderValues.BasicThinLines), Size = 12 }
                                )
                            );
                            table.AppendChild(tblProperties);
                            #region Header table
                            TableRow rowHeader = new TableRow();
                            TableCell cell1 = new TableCell(new Paragraph(new Run(new Text($"TT"))));
                            TableCell cell2 = new TableCell(new Paragraph(new Run(new Text($"Đơn vị, chỉ tiêu, nội dung kiến nghị"))));
                            TableCell cell3 = new TableCell(new Paragraph(new Run(new Text($"Số tiền"))));
                            TableCell cell4 = new TableCell(new Paragraph(new Run(new Text($"Thuyết minh nguyên nhân"))));
                            TableCell cell5 = new TableCell(new Paragraph(new Run(new Text($"Kết quả thực hiện (Số tiền)"))));
                            TableCell cell6 = new TableCell(new Paragraph(new Run(new Text($"Tài liệu/ Hồ sơ minh chứng"))));
                            rowHeader.Append(cell1);
                            rowHeader.Append(cell2);
                            rowHeader.Append(cell3);
                            rowHeader.Append(cell4);
                            rowHeader.Append(cell5);
                            rowHeader.Append(cell6);
                            table.Append(rowHeader);
                            #endregion

                            #region Gendata table
                            for (int i = 0; i < 3; i++)
                            {
                                TableRow row = new TableRow();
                                for (int j = 0; j < 6; j++)
                                {
                                    TableCell cell = new TableCell(new Paragraph(new Run(new Text($"Cell {i + 1},{j + 1}"))));
                                    row.Append(cell);
                                }
                                table.Append(row);
                            }
                            #endregion
                            paragraph.Parent.InsertAfter(table, paragraph);
                            paragraph.Remove();
                        }
                    }
                    #endregion
                }
                foreach (var paragraph in body.Descendants<Paragraph>())
                {
                    foreach (var run in paragraph.Descendants<Run>())
                    {
                        RunProperties runProperties = run.RunProperties ?? new RunProperties();
                        RunFonts runFonts = new RunFonts { Ascii = "Time New Roman" };
                        runProperties.Append(runFonts);
                        run.RunProperties = runProperties;
                    }
                }
            }
        }

        public async Task SaveReportTemplate(string fullPath, string auditValue, string yearValue, string fileOldName, string fileExt, string fileSize, string fileName, string netWorkPath)
        {
            var templateId = Guid.NewGuid().ToString();
            var template = new TblMdTemplateReport
            {
                Id = templateId,
                Year = yearValue,
                AuditPeriod = auditValue,
                FileName = fileName,
                FileOldName = fileOldName,
                FileExt = fileExt,
                FileSize = fileSize,
                FullPath = fullPath,
                NetworkPath = netWorkPath,
            };
            _dbContext.tblMdTemplateReport.Add(template);

            List<string> lstTextElement = new List<string>();
            WordDocumentService wordDocumentService = new WordDocumentService();
            using (WordprocessingDocument doc = WordprocessingDocument.Open(fullPath, true))
            {
                lstTextElement = wordDocumentService.FindTextElement(doc);
                if (lstTextElement.Count() > 0)
                {
                    foreach (var e in lstTextElement)
                    {
                        var element = new TblMdTemplateReportElement
                        {
                            Id = Guid.NewGuid().ToString(),
                            FileId = templateId,
                            TextElement = e,
                        };
                        _dbContext.tblMdTemplateReportElement.Add(element);
                    }
                }
            }
            await _dbContext.SaveChangesAsync();
        }

        public async Task<OrganizeDto> GetListOrg(string fileId, string textElement)
        {
            try
            {
                var lstNode = new List<OrganizeDto>();
                var rootNode = new OrganizeDto()
                {
                    Id = "ORG",
                    PId = "-",
                    Name = "1.1 Sở Tài Chính",
                    Title = "1.1 Sở Tài Chính",
                    Key = "OPL"
                };
                lstNode.Add(rootNode);
                var Orgs = _dbContext.tblMdTemplateReportMapping.Where(x => x.FileId == fileId && x.TextElement == textElement).Select(x => x.OrgCode).Distinct().ToList();
                var orgDetails = _dbContext.tblAdOrganize.ToList();
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
                        Checked = Orgs.Contains(org.Id) ? true : false,
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
                this.Status = false;
                this.Exception = ex;
                return null;
            }
        }
        public async Task<OpinionListDto> GetListOpinion(string fileId, string textElement, string OrgCode, string timeYear, string auditPeriod)
        {
            try
            {

                var mgCode = _dbContext.tblMdMgOpinionList.Where(x => x.TimeYear == timeYear && x.AuditPeriod == auditPeriod).Select(x => x.Code).FirstOrDefault();
                var lstNode = new List<OpinionListDto>();
                var rootNode = new OpinionListDto()
                {
                    Id = "OPL",
                    PId = "-",
                    Name = "OPL. DANH SÁCH KIẾN NGHỊ",
                    Title = "OPL. DANH SÁCH KIẾN NGHỊ",
                    Key = "OPL"
                };
                lstNode.Add(rootNode);
                var Opls = _dbContext.tblMdTemplateReportMapping.Where(x => x.FileId == fileId && x.TextElement == textElement && x.OrgCode == OrgCode).Select(x => x.OpinionCode).Distinct().ToList();
                //var oplDetails = _dbContext.TblBuOpinionLists.Distinct().ToList();
                var oplDetails = (await _serviceOpinion.GetAll()).Where(x => x.MgCode == mgCode).OrderBy(x => x.OrderNumber).ToList();
                foreach (var opl in oplDetails)
                {

                    var node = new OpinionListDto()
                    {
                        Code = opl.Code,
                        Id = opl.Id,
                        Name = opl.Name,
                        PId = rootNode.Id,
                        OrderNumber = opl.OrderNumber,
                        Title = $"{opl.Id} _ {opl.Name}",
                        Key = opl.Id,
                        Checked = Opls.Contains(opl.Code.ToString().ToUpperInvariant()) ? true : false,
                    };
                    lstNode.Add(node);
                }
                var nodeDict = lstNode.ToDictionary(n => n.Id);
                foreach (var item in lstNode.Where(n => n.Id != "ORG"))
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
                this.Status = false;
                this.Exception = ex;
                return null;
            }

        }


        public async Task SaveTemplateReport(TemplateReportMappingDto dto)
        {
            var typeElement = dto.Type;
            if (typeElement == "1")
            {
                var template = new TblMdTemplateReportMapping
                {
                    Id = Guid.NewGuid().ToString(),
                    FileId = dto.FileId,
                    TextElement = dto.TextElement,
                    Type = dto.Type,
                    ValueInput = dto.ValueInput,
                    IsActive = true
                };
                _dbContext.tblMdTemplateReportMapping.Add(template);
                await _dbContext.SaveChangesAsync();
            }
            else if (typeElement == "2")
            {
                var listOpinion = dto.OpinionCode.ToList();
                foreach (var item in listOpinion)
                {
                    var template = new TblMdTemplateReportMapping
                    {
                        Id = Guid.NewGuid().ToString(),
                        FileId = dto.FileId,
                        TextElement = dto.TextElement,
                        Type = dto.Type,
                        OrgCode = dto.OrgCode,
                        OpinionCode = item.ToUpperInvariant(),
                        IsActive = true
                    };
                    _dbContext.tblMdTemplateReportMapping.Add(template);

                }
                await _dbContext.SaveChangesAsync();
            }
            else
            {
                //var listTableCode = dto.TableCode.ToList();
                //foreach (var item in listTableCode)
                //{
                var template = new TblMdTemplateReportMapping
                {
                    Id = Guid.NewGuid().ToString(),
                    FileId = dto.FileId,
                    TextElement = dto.TextElement,
                    Type = dto.Type,
                    OrgCode = dto.OrgCode,
                    TableCode = dto.TableCode,
                    IsActive = true
                };
                _dbContext.tblMdTemplateReportMapping.Add(template);
                //}
                await _dbContext.SaveChangesAsync();
            }

        }
        public async Task<TblMdTemplateReport> GetTemplate(string id, string year, string audit)
        {
            try
            {
                WordDocumentService wordDocumentService = new WordDocumentService();
                var templatePath = _dbContext.tblMdTemplateReport.Where(x => x.Id == id).Select(x => x.FullPath).FirstOrDefault();
                var curPath = _dbContext.tblMdTemplateReport.Where(x => x.Id == id).Select(x => x.NetworkPath).FirstOrDefault();
                var lstTextElement = _dbContext.tblMdTemplateReportElement.Where(x => x.FileId == id).Select(x => x.TextElement).ToList();

                var header = _dbContext.tblMdListAudit.Where(x => x.TimeYear == year && x.AuditPeriod == audit).FirstOrDefault().Code;
                var dataOpinion = _dbContext.tblBuOpinionDetail.Where(x => x.MgCode == header);

                var fileNameNew = Guid.NewGuid().ToString() + ".docx";

                var documentPath = Path.Combine(curPath, fileNameNew);

                using (var template = File.OpenRead(templatePath))
                using (var documentStream = File.Open(documentPath, FileMode.OpenOrCreate))
                {
                    template.CopyTo(documentStream);

                    using (var document = WordprocessingDocument.Open(documentStream, true))
                    {
                        document.MainDocumentPart.Document.Save();
                    }
                }
                using (WordprocessingDocument doc = WordprocessingDocument.Open(documentPath, true))
                {
                    MainDocumentPart mainPart = doc.MainDocumentPart;
                    Body body = mainPart.Document.Body;


                    foreach (var text in lstTextElement.Distinct())
                    {
                        var check = _dbContext.tblMdTemplateReportMapping.Where(x => x.FileId == id && x.TextElement == text);
                        if (check == null || check.Count() == 0) continue;
                        if (check.FirstOrDefault().Type == "1")
                        {

                        }
                        else if (check.FirstOrDefault().Type == "2")
                        {
                            var fillText = "";
                            var order = 1;
                            foreach (var c in check)
                            {
                                var orgCode = c.OrgCode;
                                var opi = (await _dbContext.TblBuOpinionLists.FirstOrDefaultAsync(x => x.Code.ToString() == c.OpinionCode))?.Id;
                                var opiTxt = (await _dbContext.TblBuOpinionLists.FirstOrDefaultAsync(x => x.Code.ToString() == c.OpinionCode))?.Name;
                                var i = dataOpinion.Where(x => x.OrgCode == orgCode && x.OpinionCode == opi.ToString()).FirstOrDefault();
                                if (i == null) continue;
                                fillText += $"{order}. {opiTxt} new_line {i.ContentReport}";
                                order++;
                            }
                            wordDocumentService.ReplaceStringInWordDocumennt(doc, text, fillText);
                        }
                        else if (check.FirstOrDefault().Type == "3")
                        {

                        }



                        //var textReplace = text.Replace("##", "").Replace("@@", "").Split("_").ToList();

                        //#region Gen kiến nghị
                        //if (textReplace.Count() == 3)
                        //{
                        //    var dataTreeOpinion = _dbContext.TblBuOpinionLists.Where(x => x.MgCode == headerData.OpinionCode).ToList();
                        //    var nameOpinion = dataTreeOpinion.Where(x => x.Id == textReplace[1]).FirstOrDefault().Name;
                        //    var dataFill = dataDetail.Where(x => x.OrgCode == textReplace[0] && x.OpinionCode == textReplace[1]).FirstOrDefault();
                        //    if (textReplace[2] == "N")
                        //    {
                        //        wordDocumentService.ReplaceStringInWordDocumennt(doc, text, nameOpinion);
                        //    }
                        //    else
                        //    {
                        //        wordDocumentService.ReplaceStringInWordDocumennt(doc, text, dataFill.ContentReport);
                        //    }
                        //}
                        //#endregion

                    }
                    foreach (var paragraph in body.Descendants<Paragraph>())
                    {
                        foreach (var run in paragraph.Descendants<Run>())
                        {
                            RunProperties runProperties = run.RunProperties ?? new RunProperties();
                            RunFonts runFonts = new RunFonts { Ascii = "Time New Roman" };
                            runProperties.Append(runFonts);
                            run.RunProperties = runProperties;
                        }
                    }
                }
                using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(documentPath, true))
                {
                    MainDocumentPart mainPart = wordDocument.MainDocumentPart;
                    var paragraphs = mainPart.Document.Body.Elements<Paragraph>();

                    foreach (var paragraph in paragraphs)
                    {
                        foreach (var run in paragraph.Elements<Run>())
                        {
                            foreach (var text in run.Elements<Text>())
                            {
                                if (text.Text.Contains("new_line"))
                                {
                                    var parentRun = text.Parent;
                                    text.Remove();

                                    // Create a new run with the new line text and a break
                                    Run newRun = new Run();
                                    newRun.AppendChild(new Text("new line text"));
                                    newRun.AppendChild(new Break());

                                    parentRun.InsertAfter(newRun, parentRun.LastChild);
                                }
                            }
                        }
                    }

                    mainPart.Document.Save();
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
    }
}
