using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelFlow;

internal static class ExcelExtensions
{
    /// <summary>
    /// Style index (into cellXfs) applied by generated writers to date-only DateTime values.
    /// </summary>
    internal const uint DateStyleIndex = 1;

    /// <summary>
    /// Style index (into cellXfs) applied by generated writers to DateTime values with a time component.
    /// </summary>
    internal const uint DateTimeStyleIndex = 2;

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

        AddMinimalStylesheet(workbookPart);

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

    /// <summary>
    /// Adds a minimal stylesheet so that DateTime cells (written as OADate numbers)
    /// are displayed as dates instead of raw numbers.
    /// cellXfs: 0 = default, 1 = built-in date format (14), 2 = built-in date-time format (22).
    /// </summary>
    private static void AddMinimalStylesheet(WorkbookPart workbookPart)
    {
        WorkbookStylesPart stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();

        stylesPart.Stylesheet = new Stylesheet(
            new Fonts(new Font()) { Count = 1U },
            new Fills(
                new Fill(new PatternFill { PatternType = PatternValues.None }),
                new Fill(new PatternFill { PatternType = PatternValues.Gray125 })) { Count = 2U },
            new Borders(new Border()) { Count = 1U },
            new CellStyleFormats(new CellFormat()) { Count = 1U },
            new CellFormats(
                new CellFormat(),
                new CellFormat { NumberFormatId = 14U, FontId = 0U, FillId = 0U, BorderId = 0U, FormatId = 0U, ApplyNumberFormat = true },
                new CellFormat { NumberFormatId = 22U, FontId = 0U, FillId = 0U, BorderId = 0U, FormatId = 0U, ApplyNumberFormat = true }) { Count = 3U });

        stylesPart.Stylesheet.Save();
    }
}
