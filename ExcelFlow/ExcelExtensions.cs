using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelFlow;

internal static class ExcelExtensions
{
    public static void ToExcel<T>(this IEnumerable<T> data, string filePath, string sheetName) where T : class
    {
        using FileStream stream = File.Open(filePath, FileMode.Create, FileAccess.ReadWrite);
        
        data.ToExcel(stream, sheetName);
    }

    public static void ToExcel<T>(this IEnumerable<T> data, Stream stream, string sheetName) where T : class
    {
        if (data == null) throw new ArgumentNullException(nameof(data));
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (string.IsNullOrEmpty(sheetName)) sheetName = "Sheet1";

        PropertyInfo[] properties = typeof(T).GetProperties();
        List<ExportColumnMap<T>> exportMap = new List<ExportColumnMap<T>>();

        foreach (PropertyInfo prop in properties)
        {
            if (!prop.CanRead) 
                continue;
            
            ExcelColumnAttribute? attr = prop.GetCustomAttribute<ExcelColumnAttribute>();
            string expectedName = attr?.Name ?? prop.Name;
            
            Func<T, object> getter = ExpressionCompiler.CompileGetter<T>(prop);

            exportMap.Add(new ExportColumnMap<T>(expectedName, prop.PropertyType, getter));
        }

        using SpreadsheetDocument document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
        WorkbookPart workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        
        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

        using (OpenXmlWriter writer = OpenXmlWriter.Create(worksheetPart))
        {
            writer.WriteStartElement(new Worksheet());
            writer.WriteStartElement(new SheetData());

            writer.WriteStartElement(new Row());
            foreach (ExportColumnMap<T> col in exportMap)
            {
                writer.WriteElement(new Cell
                {
                    CellValue = new CellValue(col.ColumnName),
                    DataType = CellValues.String
                });
            }
            writer.WriteEndElement();
            
            foreach (T item in data)
            {
                writer.WriteStartElement(new Row());
                foreach (ExportColumnMap<T> col in exportMap)
                {
                    object? value = col.Getter(item);
                    Cell cell = new Cell();

                    if (value is null)
                    {
                        // Empty cell for nullable
                    }
                    else if (value is int || value is long || value is decimal || value is double || value is float)
                    {
                        cell.CellValue = new CellValue(Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty);
                        cell.DataType = CellValues.Number;
                    }
                    else if (value is DateTime dateTime)
                    {
                        cell.CellValue = new CellValue(dateTime.ToString("yyyy-MM-ddTHH:mm:ss"));
                        cell.DataType = CellValues.String; 
                    }
                    else if (value is bool boolean)
                    {
                        cell.CellValue = new CellValue(boolean ? "1" : "0");
                        cell.DataType = CellValues.Boolean;
                    }
                    else
                    {
                        cell.CellValue = new CellValue(value.ToString() ?? string.Empty);
                        cell.DataType = CellValues.String;
                    }

                    writer.WriteElement(cell);
                }
                writer.WriteEndElement();
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

    public static (object?, bool) SafeConvert(object? value, Type targetType)
    {
        if (value == null || value == DBNull.Value || string.IsNullOrWhiteSpace(value.ToString()))
        {
            if (Nullable.GetUnderlyingType(targetType) != null || !targetType.IsValueType)
                return (null, true);
            
            return (Activator.CreateInstance(targetType), true);
        }
        
        Type underlyingType = Nullable.GetUnderlyingType(targetType) ?? targetType;
        
        if (underlyingType == typeof(DateTime))
        {
            if (value is DateTime dt) 
                return (dt, true);

            if (value is double d) 
                return (DateTime.FromOADate(d), true);

            if (value is string strDate)
            {
                if (DateTime.TryParse(strDate, out DateTime parsedDt)) 
                    return (parsedDt, true);
            
                if (double.TryParse(strDate.Replace(',', '.'), 
                        System.Globalization.NumberStyles.Any, 
                        System.Globalization.CultureInfo.InvariantCulture, out double oaDate))
                {
                    return (DateTime.FromOADate(oaDate), true);
                }
            }
        }

        try
        {
            if (underlyingType == typeof(decimal) || underlyingType == typeof(double) ||
                underlyingType == typeof(float))
            {
                if (value is string strValue)
                {
                    strValue = strValue.Replace(',', '.');
                    return (Convert.ChangeType(strValue, underlyingType,
                        System.Globalization.CultureInfo.InvariantCulture), true);
                }
            }

            return (Convert.ChangeType(value, underlyingType, System.Globalization.CultureInfo.InvariantCulture), true);
        }
        catch
        {
            if (Nullable.GetUnderlyingType(targetType) != null || !targetType.IsValueType)
                return (null, false);

            return (Activator.CreateInstance(targetType), false);
        }
    }
}