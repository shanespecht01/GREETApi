
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Runtime.InteropServices;

namespace GREETApi.Services
{
    public interface IExcelService
    {
        void SendToGREET1(string filePath, string sheetName, dynamic test);
        IEnumerable<object[]> GetFromGREET1(string filePath, string sheetName);
    }

    public class ExcelService : IExcelService
    {
        public void SendToGREET1(string filePath, string sheetName, dynamic data)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, true))
            {
                var workbookPart = document.WorkbookPart;
                var workbook = workbookPart.Workbook;

                var definedNames = workbookPart.Workbook.DefinedNames;
                if (definedNames == null)
                    throw new InvalidOperationException("No defined names in the workbook.");

                Type type = data.GetType();

                // Get all properties of the object
                PropertyInfo[] properties = type.GetProperties();

                // Iterate through each property
                foreach (var property in properties)
                {
                    string rangeName = property.Name;
                    object newValue = property.GetValue(data);

                    var definedName = definedNames.Elements<DefinedName>()
                        .FirstOrDefault(d => d.Name.Value.Equals(rangeName, StringComparison.OrdinalIgnoreCase));

                    if (definedName != null)
                    {
                        var reference = definedName.Text;
                        var worksheetPart = GetWorksheetPartByReference(workbookPart, reference);
                        var cellReference = reference.Split('!').Last().Replace("$", string.Empty);

                        var worksheet = worksheetPart.Worksheet;
                        var cell = GetCell(worksheet, cellReference);

                        cell.CellValue = new CellValue(newValue.ToString());
                        cell.DataType = new EnumValue<CellValues>(CellValues.String);
                    }
                }

                workbook.Save();
            }

        }

        public IEnumerable<object[]> GetFromGREET1(string filePath, string sheetName)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, false))
            {
                var workbookPart = document.WorkbookPart;
                var workbook = workbookPart.Workbook;

                // Get defined names
                var definedNames = workbook.DefinedNames;
                if (definedNames == null)
                    throw new ArgumentException("No defined names in the workbook.");

                // Find the defined name for "Aviation_Module_Results"
                var definedName = definedNames.Elements<DefinedName>()
                    .FirstOrDefault(d => d.Name.Value.Equals("Aviation_Module_Results", StringComparison.OrdinalIgnoreCase));

                if (definedName == null)
                    throw new ArgumentException("Defined name 'Aviation_Module_Results' not found in the workbook.");

                // Get the sheet and cell reference from the defined name
                var reference = definedName.Text;
                var parts = reference.Split('!');
                var refSheetName = parts[0].Trim('\'');
                var cellReference = parts[1].Replace("$", string.Empty);

                if (!refSheetName.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                    throw new ArgumentException($"Defined name 'Aviation_Module_Results' does not refer to sheet '{sheetName}'.");

                var worksheetPart = GetWorksheetPartBySheetName(workbookPart, refSheetName);
                var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                // Calculate start and end rows for each range
                uint baseRowIndex = GetRowIndex(cellReference);
                uint feedS = baseRowIndex + 1;
                uint feedE = feedS + 26;
                uint fuelS = feedE + 2;
                uint fuelE = fuelS + 26;
                uint combS = fuelE + 2;
                uint combE = combS + 26;
                uint wtws = combE + 2;
                uint wtwe = wtws + 26;

                var data = new List<object[]>();

                // Extract data for each range
                data.AddRange(ExtractRangeData(sheetData.Elements<Row>(), feedS, feedE));
                data.AddRange(ExtractRangeData(sheetData.Elements<Row>(), fuelS, fuelE));
                data.AddRange(ExtractRangeData(sheetData.Elements<Row>(), combS, combE));
                data.AddRange(ExtractRangeData(sheetData.Elements<Row>(), wtws, wtwe));

                return data;
            }
        }

        private WorksheetPart GetWorksheetPartByReference(WorkbookPart workbookPart, string reference)
        {
            var sheetName = reference.Split('!').First().Replace("'", string.Empty);
            var sheet = workbookPart.Workbook.Sheets.Elements<Sheet>()
                .FirstOrDefault(s => s.Name.Value.Equals(sheetName, StringComparison.OrdinalIgnoreCase));

            if (sheet == null)
                throw new ArgumentException($"Sheet '{sheetName}' not found.");

            return (WorksheetPart)workbookPart.GetPartById(sheet.Id);
        }

        private WorksheetPart GetWorksheetPartBySheetName(WorkbookPart workbookPart, string sheetName)
        {
            var sheet = workbookPart.Workbook.Descendants<Sheet>()
                .FirstOrDefault(s => s.Name.Value.Equals(sheetName, StringComparison.OrdinalIgnoreCase));

            if (sheet == null)
                throw new ArgumentException($"Sheet '{sheetName}' not found.");

            return (WorksheetPart)workbookPart.GetPartById(sheet.Id);
        }

        private Cell GetCell(DocumentFormat.OpenXml.Spreadsheet.Worksheet worksheet, string cellReference)
        {
            var sheetData = worksheet.GetFirstChild<SheetData>();
            var rowNumber = GetRowIndex(cellReference);
            var columnName = GetColumnName(cellReference);

            var row = sheetData.Elements<Row>()
                .FirstOrDefault(r => r.RowIndex == rowNumber) ?? sheetData.AppendChild(new Row() { RowIndex = rowNumber });

            var cell = row.Elements<Cell>()
                .FirstOrDefault(c => string.Compare(c.CellReference.Value, cellReference, true) == 0);

            if (cell == null)
            {
                cell = new Cell() { CellReference = cellReference };
                row.AppendChild(cell);
            }

            return cell;
        }

        private uint GetRowIndex(string cellReference)
        {
            var rowPart = new string(cellReference.Where(char.IsDigit).ToArray());
            return uint.Parse(rowPart);
        }

        private string GetColumnName(string cellReference)
        {
            return new string(cellReference.Where(char.IsLetter).ToArray());
        }

        private IEnumerable<object[]> ExtractRangeData(IEnumerable<Row> rows, uint startRow, uint endRow)
        {
            var data = new List<object[]>();

            foreach (var row in rows)
            {
                uint rowIndex = row.RowIndex.Value;
                if (rowIndex >= startRow && rowIndex <= endRow)
                {
                    var cells = row.Elements<Cell>()
                        .Select(c => c.CellValue?.Text ?? string.Empty)
                        .ToArray();

                    data.Add(cells);
                }
            }

            return data;
        }
        
    }
}