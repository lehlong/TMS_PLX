using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMS.BUSINESS.Extentions
{
    public static class ExcelNPOIExtention
    {
        public static void SetCellValue(IRow row, int columnIndex, object value, ICellStyle style)
        {
            var cell = row.CreateCell(columnIndex);
            cell.CellStyle = style;

            if (value is null)
            {
                cell.SetCellValue("");
            }
            else if (double.TryParse(value.ToString(), out double number))
            {
                cell.SetCellValue(number);
            }
            else
            {
                cell.SetCellValue(value.ToString());
            }
        }
        public static ICellStyle SetCellStyleText(IWorkbook workbook)
        {
            ICellStyle style = workbook.CreateCellStyle();
            IFont font = workbook.CreateFont();
            font.FontName = "Times New Roman";
            font.FontHeightInPoints = 12;
            style.SetFont(font);
            style.Alignment = HorizontalAlignment.Left;
            style.BorderTop = BorderStyle.Thin;
            style.BorderBottom = BorderStyle.Thin;
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            return style;
        }

        public static ICellStyle SetCellStyleTextBold(IWorkbook workbook)
        {
            ICellStyle style = workbook.CreateCellStyle();
            IFont font = workbook.CreateFont();
            font.FontName = "Times New Roman";
            font.FontHeightInPoints = 12;
            font.IsBold = true;
            style.SetFont(font);
            style.Alignment = HorizontalAlignment.Left;
            style.BorderTop = BorderStyle.Thin;
            style.BorderBottom = BorderStyle.Thin;
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            return style;
        }

        public static ICellStyle SetCellStyleNumber(IWorkbook workbook)
        {
            ICellStyle style = workbook.CreateCellStyle();
            IFont font = workbook.CreateFont();
            font.FontName = "Times New Roman";
            font.FontHeightInPoints = 12;
            style.SetFont(font);
            style.Alignment = HorizontalAlignment.Right;
            style.BorderTop = BorderStyle.Thin;
            style.BorderBottom = BorderStyle.Thin;
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.DataFormat = workbook.CreateDataFormat().GetFormat("#,##0");
            return style;
        }

        public static ICellStyle SetCellStyleNumberBold(IWorkbook workbook)
        {
            ICellStyle style = workbook.CreateCellStyle();
            IFont font = workbook.CreateFont();
            font.FontName = "Times New Roman";
            font.FontHeightInPoints = 12;
            font.IsBold = true;
            style.SetFont(font);
            style.Alignment = HorizontalAlignment.Right;
            style.BorderTop = BorderStyle.Thin;
            style.BorderBottom = BorderStyle.Thin;
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.DataFormat = workbook.CreateDataFormat().GetFormat("#,##0");
            return style;
        }
    }
}
