using NPOI.SS.UserModel;

public class ExcelNPOIExtention
{
    public ICellStyle SetCellStyleText(IWorkbook workbook)
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

    public ICellStyle SetCellStyleTextBold(IWorkbook workbook)
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

    public ICellStyle SetCellStyleNumber(IWorkbook workbook)
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

    public ICellStyle SetCellStyleNumberBold(IWorkbook workbook)
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
