using Aspose.Cells;

namespace DBDoc.Doc.Excel
{
    public static class AsposeCellHelper
    {
        public static Workbook ReadExcel(string filePath)
        {
            var fileBytes = File.ReadAllBytes(filePath);
            using (var stream = new MemoryStream(fileBytes))
            {
                stream.Position = 0;
                return new Workbook(stream);
            }
        }

        public static Workbook ReadExcel(byte[] fileBytes)
        {
            using (var stream = new MemoryStream(fileBytes))
            {
                return new Workbook(stream);
            }
        }

        public static Workbook ReadExcel(Stream stream)
        {
            stream.Position = 0;
            return new Workbook(stream);
        }

        public static Workbook CreateExcel()
        {
            var excel = new Workbook();
            excel.Worksheets.Clear();
            return excel;
        }

        public static byte[] ToBytes(this Workbook workbook)
        {
            using (var stream = workbook.SaveToStream())
            {
                return stream.ToArray();
            }
        }

        public static void EvaluateAllFormulaCells(this Workbook workbook)
        {
            workbook.CalculateFormula(true);
        }

        public static Worksheet CreateSheet(this Workbook workbook, string name)
        {
            return workbook.Worksheets.Add(name);
        }

        public static Worksheet GetSheet(this Workbook workbook, string name)
        {
            return workbook.Worksheets[name];
        }

        public static Worksheet GetSheetAt(this Workbook workbook, int index)
        {
            return workbook.Worksheets[index];
        }

        public static Worksheet ReadExcelSheet(string filePath, int index)
        {
            return ReadExcel(filePath).GetSheetAt(index);
        }

        public static Worksheet ReadExcelSheet(string filePath, string name)
        {
            return ReadExcel(filePath).GetSheet(name);
        }

        public static Worksheet ReadExcelSheet(Stream stream, string name)
        {
            return ReadExcel(stream).GetSheet(name);
        }

        public static Row GetRow(this Worksheet sheet, int row)
        {
            return sheet.Cells.GetRow(row);
        }

        public static Cell GetCell(this Worksheet sheet, int row, int cell)
        {
            return sheet.Cells.GetCell(row, cell);
        }

        public static Cell GetCell(this Worksheet sheet, string cellRef)
        {
            return sheet.Cells[cellRef];
        }

        public static object GetData(this Cell cell, bool calculate = true)
        {
            if (cell == null)
            {
                return null;
            }
            if (calculate)
            {
                return cell.Value;
            }
            return cell.Formula;
        }


        public static Cell CreateCell(this Worksheet sheet, string cellRef)
        {
            return sheet.Cells[cellRef];
        }

        public static Cell CreateCell(this Worksheet sheet, int row, int col)
        {
            return sheet.Cells[row, col];
        }

        public static Aspose.Cells.Range CreateRange(this Worksheet sheet, int firstRow, int firstColumn, int totalRows, int totalColumns)
        {
            sheet.Cells.Merge(firstRow, firstColumn, totalRows, totalColumns);
            return sheet.Cells[firstRow, firstColumn].GetMergedRange();
        }


        public static Cell GetOrCreateCell(this Worksheet sheet, int row, int cell)
        {
            return sheet.GetCell(row, cell) ?? sheet.CreateCell(row, cell);
        }


        public static Cell GetOrCreateCell(this Worksheet sheet, string cellRef)
        {
            return sheet.GetCell(cellRef) ?? sheet.CreateCell(cellRef);
        }

        public static Cell SetCell(this Worksheet sheet, string cellRef, bool data)
        {
            var cell = sheet.GetOrCreateCell(cellRef);
            cell.Value = data;
            return cell;
        }

        public static Column SetColumnWidth(this Worksheet sheet, int columnIndex, double width)
        {
            var column = sheet.Cells.Columns[columnIndex];
            column.Width = width;
            return column;
        }

        public static Cell SetCell(this Worksheet sheet, string cellRef, double data)
        {
            var cell = sheet.GetOrCreateCell(cellRef);
            cell.Value = data;
            return cell;
        }
        public static Cell SetCell(this Worksheet sheet, string cellRef, string data)
        {
            var cell = sheet.GetOrCreateCell(cellRef);
            cell.Value = data;
            return cell;
        }

        public static int GetIndex(this Row row, string text, int defaultIndex = -1)
        {
            for (int i = 0; i < row.LastCell.Column; i++)
            {
                if (row.GetCellOrNull(i)?.StringValue == text)
                {
                    return i;
                }
            }
            return defaultIndex;
        }

        public static Aspose.Cells.Range SetValue(this Aspose.Cells.Range range, string data)
        {
            range.Value = data;
            return range;
        }

        public static Cell SetValue(this Cell cell, string data)
        {
            cell.Value = data;
            return cell;
        }

        public static Cell SetValue(this Cell cell, double data)
        {
            cell.Value = data;
            return cell;
        }

        public static Cell SetCellValue(this Cell cell, object data)
        {
            cell.Value = data;
            return cell;
        }

        public static Cell SetBorder(this Cell cell, CellBorderType borderStype = CellBorderType.Thin)
        {
            return cell.SetBorder(borderStype, borderStype, borderStype, borderStype);
        }

        public static Cell SetBorder(this Cell cell, CellBorderType borderTop, CellBorderType borderRight, CellBorderType borderBottom, CellBorderType borderLeft)
        {
            var cellStyle = cell.Worksheet.Workbook.CreateStyle();
            cellStyle.Borders[BorderType.TopBorder].LineStyle = borderTop;
            cellStyle.Borders[BorderType.RightBorder].LineStyle = borderRight;
            cellStyle.Borders[BorderType.BottomBorder].LineStyle = borderBottom;
            cellStyle.Borders[BorderType.LeftBorder].LineStyle = borderLeft;
            cell.SetStyle(cellStyle);
            return cell;
        }

        public static Cell SetCellStyle(this Cell cell, Style cellStyle)
        {
            cell.SetStyle(cellStyle);
            return cell;
        }
    }
}
