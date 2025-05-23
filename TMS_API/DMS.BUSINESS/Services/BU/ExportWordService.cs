using System.Text.RegularExpressions;
using System.Text;
using DMS.BUSINESS.Models;
using DMS.CORE;
using DMS.CORE.Entities.MD;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.EntityFrameworkCore;
using DocumentFormat.OpenXml.Packaging;
using DMS.CORE.Entities.BU;

namespace DMS.BUSINESS.Services.BU
{
    public class ExportWordService
    {
        private readonly AppDbContext _dbContext;
        public ExportWordService(AppDbContext dbContext) => _dbContext = dbContext;

        public async Task<string> GenarateWord(List<CustomBBDOExportWord> lstCustomerChecked, string headerId, CalculateDiscountOutputModel data)
        {
            try
            {
                var folderPath = Path.Combine($"Uploads/Word/{DateTime.Now.ToString("yyyy/MM/dd")}");
                if (!Directory.Exists(folderPath))
                    Directory.CreateDirectory(folderPath);

                var fileName = $"ThongBaoGia_{DateTime.Now:ddMMyyyy_HHmmss}.docx";
                var templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template", "ThongBaoGia.docx");
                var fullPath = Path.Combine(Directory.GetCurrentDirectory(), folderPath, fileName);
                if (!File.Exists(fullPath))
                    File.Copy(templatePath, fullPath, true);
                TblBuCalculateDiscount oldHeader = null;
                var header = await _dbContext.TblBuCalculateDiscount.FindAsync(headerId);
                if (header.QuyetDinhSo != null)
                {
                    var oldHeaders = await _dbContext.TblBuCalculateDiscount
                        //.Where(x => x.QuyetDinhSo == header.QuyetDinhSo)
                        .Where(x => x.QuyetDinhSo == header.QuyetDinhSo && x.Id != headerId)
                        .ToListAsync();
                    if(oldHeader != null)
                    {
                        oldHeader = oldHeaders
                            .OrderBy(x => x.Date)
                            .FirstOrDefault();
                    }
                    else
                    {
                        oldHeader = header;
                    }
                }
                else
                {
                    oldHeader = header;
                }
                var signer = await _dbContext.TblMdSigner.FirstOrDefaultAsync(x => x.Code == header.SignerCode);
                var goodsList = await _dbContext.TblMdGoods.Where(x => x.IsActive == true).ToListAsync();
                var lastCode = lstCustomerChecked.LastOrDefault()?.code;
                bool isN1 = false;

                using var templateDoc = WordprocessingDocument.Open(templatePath, false);
                var templateBody = templateDoc.MainDocumentPart.Document.Body;

                using var destDoc = WordprocessingDocument.Open(fullPath, true);
                var destBody = destDoc.MainDocumentPart.Document.Body;
                destBody.RemoveAllChildren();

                for (int i = 0; i < lstCustomerChecked.Count; i++)
                {
                    var customer = lstCustomerChecked[i];
                    var tempBody = (Body)templateBody.CloneNode(true);
                    var customerData = data.Vk11Bb.Where(x => x.Col4 == customer.code).ToList();
                    var customerInfo = await _dbContext.TblBuInputCustomerBbdo.FirstOrDefaultAsync(x => x.Code == customer.code && x.HeaderId == headerId);

                    ReplaceStringInWordBody(tempBody, "##COMPANY@@", customerInfo?.Name);
                    ReplaceStringInWordBody(tempBody, "##ADDRESS@@", customerInfo?.Adrress);
                    ReplaceStringInWordBody(tempBody, "##CHIPHI@@", customerInfo.DeliveryGroupCode != "N1" ? "và chi phí vân chuyển" : "");

                    AppendCustomerTable(tempBody, customer, data, goodsList, customerData, out isN1);

                    if (i == 0)
                    {
                        foreach (var element in tempBody.Elements())
                        {
                            destBody.Append(element.CloneNode(true));
                        }
                    }
                    else
                    {
                        destBody.AppendChild(new Paragraph(new Run(new Break { Type = BreakValues.Page })));
                        foreach (var element in tempBody.Elements())
                        {
                            destBody.Append(element.CloneNode(true));
                        }
                    }
                }
                //ngày 24 tháng 10 năm 2024
                var replacements = new Dictionary<string, string>
                {
                    ["##DATE@@"] = $"{header?.Date.Hour:D2}h00 ngày {header?.Date.Day:D2} tháng {header?.Date.Month:D2} năm {header?.Date.Year}",
                    ["##DATE3@@"] = $"ngày {header?.Date.Day:D2} tháng {header?.Date.Month:D2} năm {header?.Date.Year}",
                    ["##DATE2@@"] = $"ngày {oldHeader?.Date.Day:D2} tháng {oldHeader?.Date.Month:D2} năm {oldHeader?.Date.Year}",
                    ["##QUYET_DINH_SO@@"] = header?.QuyetDinhSo ?? "",
                    ["##DAI_DIEN@@"] = signer?.Code != "TongGiamDoc" ? "KT.GIÁM ĐỐC CÔNG TY" : "",
                    ["##NGUOI_DAI_DIEN@@"] = signer.Position,
                    ["##TEN@@"] = signer.Name,
                };

                foreach (var (key, value) in replacements)
                    ReplaceStringInWordBody(destBody, key, value);

                destDoc.MainDocumentPart.Document.Save();
                return $"{folderPath}/{fileName}";
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                return null;
            }
        }


        public Body ReplaceStringInWordBody(Body body, string search, string replace)
        {
            if (body == null || string.IsNullOrEmpty(search))
                return body;

            var texts = body.Descendants<Text>().Where(t => !string.IsNullOrEmpty(t.Text)).ToList();

            if (texts.Count == 0)
                return body;

            var fullText = new StringBuilder(texts.Sum(t => t.Text.Length));
            var offsets = new int[texts.Count];

            for (int i = 0; i < texts.Count; i++)
            {
                offsets[i] = fullText.Length;
                fullText.Append(texts[i].Text);
            }

            var escapedSearch = Regex.Escape(search);
            var matches = Regex.Matches(fullText.ToString(), escapedSearch).Cast<Match>().Reverse().ToList();

            foreach (var match in matches)
            {
                int start = match.Index, end = match.Index + match.Length;

                int startIdx = FindTextIndex(texts, offsets, start);
                int endIdx = FindTextIndex(texts, offsets, end - 1);

                if (startIdx == endIdx)
                {
                    var text = texts[startIdx];
                    int localStart = start - offsets[startIdx];
                    text.Text = text.Text[..localStart] + replace + text.Text[(localStart + search.Length)..];
                }
                else
                {
                    var first = texts[startIdx];
                    int firstCut = start - offsets[startIdx];
                    first.Text = first.Text[..firstCut];

                    for (int j = startIdx + 1; j < endIdx; j++)
                        texts[j].Text = string.Empty;

                    var last = texts[endIdx];
                    int lastCut = end - offsets[endIdx];
                    last.Text = replace + last.Text[lastCut..];
                }
            }

            return body;
        }

        private int FindTextIndex(List<Text> texts, int[] offsets, int position)
        {
            int left = 0, right = texts.Count - 1;
            
            while (left <= right)
            {
                int mid = left + (right - left) / 2;
                int start = offsets[mid];
                int end = mid < offsets.Length - 1 ? offsets[mid + 1] : start + texts[mid].Text.Length;
                
                if (position >= start && position < end)
                    return mid;
                    
                if (position < start)
                    right = mid - 1;
                else
                    left = mid + 1;
            }
            
            return left < texts.Count ? left : texts.Count - 1;
        }

        void AppendCustomerTable(Body body, CustomBBDOExportWord customer, CalculateDiscountOutputModel data, List<TblMdGoods> goodsList, List<VK11Model> customerData, out bool isN1Flag)
        {
            isN1Flag = customer.deliveryGroupCode == "N1";
            var paragraph = body.Descendants<Paragraph>().FirstOrDefault(p => p.InnerText.Contains("##TABLE@@"));
            if (paragraph == null) return;
            var table = CreateN1Table(data, customer.code);
            //var table = isN1Flag
            //    ? CreateN1Table(data, customer.code)
            //    : CreateNormalTable(customerData, goodsList);

            paragraph.Parent.InsertAfter(table, paragraph);
            paragraph.Remove();
        }

        Table CreateN1Table(CalculateDiscountOutputModel data, string cusCode)
        {
            var table = InitTableWithBorders();
            var row1 = new TableRow();
            row1.Append(CreateHeaderCell("STT", 1, 2, 26, 200));
            row1.Append(CreateHeaderCell("Mặt hàng, quy cách", 1, 2, 26, 5500));
            row1.Append(CreateHeaderCell("Đơn vị tính", 1, 2, 26, 1500));
            row1.Append(CreateHeaderCell("Đơn giá đã có 10% VAT", 3, 1, 26, 3000));

            var row2 = new TableRow();
            row2.Append(CreateHeaderCell("", 1, -1, 26, 200));
            row2.Append(CreateHeaderCell("", 1, -1, 26, 5500));
            row2.Append(CreateHeaderCell("", 1, -1, 26, 1500));
            row2.Append(CreateHeaderCell("Giá bán lẻ Petrolimex công tại Vùng 2", 1, 1, 26, 3000));
            row2.Append(CreateHeaderCell("Chiết khấu", 1, 1, 26, 1000));
            row2.Append(CreateHeaderCell("Giá bán cho bên mua", 1, 1, 26, 1500));

            table.Append(row1, row2);

            int stt = 1;
            var lstCus = data.Bbdo.Where(x => x.CustomerCode == cusCode).ToList();
            foreach (var cus in lstCus)
            {
                var item = data.Dlg.Dlg6.Where(i => i.LocalCode == "V2" && i.GoodCode == cus.GoodCode).FirstOrDefault();
                var row = new TableRow();
                row.Append(CreateCell((stt++).ToString(), true, 26, 200));
                row.Append(CreateCell(item?.GoodName, true, 26, 5500));
                row.Append(CreateCell("Đ/lít tt", true, 26, 1500));
                row.Append(CreateCell(item.Col6.ToString("N0"), true, 26, 3000));
                row.Append(CreateCell((item.Col6 - cus.Col14).ToString("N0"), true, 26, 1000));
                row.Append(CreateCell(cus.Col14.ToString("N0"), true, 26, 1500));
                table.Append(row);
            }
            return table;
        }

        private Table CreateNormalTable(List<VK11Model> customerData, List<TblMdGoods> goodsList)
        {
            var table = InitTableWithBorders();
            var header = new TableRow();
            header.Append(CreateHeaderCell("STT", 1, 1, 26, 500));
            header.Append(CreateHeaderCell("Mặt hàng", 1, 1, 26, 4000));
            header.Append(CreateHeaderCell("Điểm giao hàng", 1, 1, 26, 6000));
            header.Append(CreateHeaderCell("Đơn giá", 2, 1, 26, 3000));
            table.Append(header);

            int stt = 1;
            foreach (var item in customerData)
            {
                var good = goodsList.FirstOrDefault(g => g.Code == item?.Col5);
                var row = new TableRow();
                row.Append(CreateCell((stt++).ToString(), false, 26, 500));
                row.Append(CreateCell(good?.Name, true, 26, 4000));
                row.Append(CreateCell(item?.Address, false, 26, 6000));
                row.Append(CreateCell(item?.Col8.ToString("N0"), true, 26, 1500));
                row.Append(CreateCell("Đ/lít tt", true, 26, 1500));
                table.Append(row);
            }
            return table;
        }

        private Table InitTableWithBorders() => new Table(
            new TableProperties(
                new TableBorders(
                    new TopBorder { Val = BorderValues.Single, Size = 4 },
                    new BottomBorder { Val = BorderValues.Single, Size = 4 },
                    new LeftBorder { Val = BorderValues.Single, Size = 4 },
                    new RightBorder { Val = BorderValues.Single, Size = 4 },
                    new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                    new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }
                )
            )
        );

        private TableCell CreateHeaderCell(string text, int gridSpan, int rowSpan, int fontSize, int width)
        {
            var cell = new TableCell(new Paragraph(new Run(new Text(text)))
            {
                ParagraphProperties = new ParagraphProperties(new Justification { Val = JustificationValues.Center })
            });

            var run = cell.Descendants<Run>().First();
            run.RunProperties = new RunProperties(new Bold(), new FontSize { Val = fontSize.ToString() });

            var props = new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = width.ToString() });
            if (gridSpan > 1) props.Append(new GridSpan { Val = gridSpan });
            if (rowSpan > 1) props.Append(new VerticalMerge { Val = MergedCellValues.Restart });
            else if (rowSpan == -1) props.Append(new VerticalMerge { Val = MergedCellValues.Continue });
            props.Append(new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center });

            cell.Append(props);
            return cell;
        }

        private TableCell CreateCell(string text, bool isBold, int fontSize, int width)
        {
            var runProps = new RunProperties(new FontSize { Val = fontSize.ToString() });
            if (isBold) runProps.Append(new Bold());

            var para = new Paragraph(new Run(runProps, new Text(text)))
            {
                ParagraphProperties = new ParagraphProperties(new Justification { Val = JustificationValues.Center })
            };

            var props = new TableCellProperties(
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = width.ToString() },
                new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center }
            );

            return new TableCell(para) { TableCellProperties = props };
        }

    }
}
