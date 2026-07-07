using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelFlow;

internal static class ExcelExtensions
{
    public static void ToExcel<T>(this IEnumerable<T> data, string filePath, string sheetName) where T : class, 
        IExcelFlowSerializable<T>, new()
    {
        using FileStream stream = File.Open(filePath, FileMode.Create, FileAccess.ReadWrite);
        
        data.ToExcel(stream, sheetName);
    }

    public static void ToExcel<T>(this IEnumerable<T> data, Stream stream, string sheetName) where T : class, 
        IExcelFlowSerializable<T>, new()
    {
        if (data == null) throw new ArgumentNullException(nameof(data));
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (string.IsNullOrEmpty(sheetName)) sheetName = "Sheet1";

        using SpreadsheetDocument document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
        WorkbookPart workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        
        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

        using (OpenXmlWriter writer = OpenXmlWriter.Create(worksheetPart))
        {
            writer.WriteStartElement(new Worksheet());
            writer.WriteStartElement(new SheetData());

            T.WriteHeaders(writer);
            
            foreach (T item in data)
            {
                T.WriteRow(writer, item);
            }

            writer.WriteEndElement();
            writer.WriteEndElement();
        }

        Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
        Sheet sheet = new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = 1,
            Name = sheetName
        };
        sheets.Append(sheet);

        workbookPart.Workbook.Save();
    }

}