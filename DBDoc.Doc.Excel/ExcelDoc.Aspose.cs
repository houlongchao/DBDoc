using DBDoc.Dto;
using Aspose.Cells;

namespace DBDoc.Doc.Excel
{
    public partial class ExcelDoc
    {
        private void ExportExcelByAspose(string fileName, DBDto dbDto)
        {
            if (File.Exists("Aspose.license"))
            {
                new License().SetLicense("Aspose.license");
            }
            else if (File.Exists("Aspose.license.base64"))
            {
                new License().SetLicense(new MemoryStream(Convert.FromBase64String(File.ReadAllText("Aspose.license.base64"))));
            }

            if (File.Exists(fileName))
            {
                File.Delete(fileName);
            }

            var excel = AsposeCellHelper.CreateExcel();

            CreateHomeSheet(excel, "修订日志", dbDto);

            CreateOverviewSheet(excel, "数据库表目录", dbDto);

            CreateTableSheet(excel, "数据库表结构", dbDto);

            CreateViewsSheet(excel, "数据库视图", dbDto);

            CreateProcsSheet(excel, "数据库存储过程", dbDto);

            excel.Save(fileName);
        }

        private Style CreateTitleStyle(Workbook workbook, int fontSize = 14)
        {
            var titleStyle = workbook.CreateStyle();
            titleStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            titleStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
            titleStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            titleStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            titleStyle.HorizontalAlignment = TextAlignmentType.Center;
            titleStyle.VerticalAlignment = TextAlignmentType.Center;
            titleStyle.Font.IsBold = true;
            titleStyle.Font.Size = fontSize;
            titleStyle.ForegroundColor = System.Drawing.Color.FromArgb(198, 217, 241);
            titleStyle.Pattern = BackgroundType.Solid;

            return titleStyle;
        }

        private Style CreateCellStyle(Workbook workbook, TextAlignmentType horizontalAlignment = TextAlignmentType.Center, string fontName = "Arial", bool shrinkToFit = false)
        {
            var cellStyle = workbook.CreateStyle();
            cellStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            cellStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
            cellStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            cellStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            cellStyle.HorizontalAlignment = horizontalAlignment;
            cellStyle.VerticalAlignment = TextAlignmentType.Center;
            cellStyle.Font.IsBold = false;
            cellStyle.Font.Size = 10;
            cellStyle.Font.Name = fontName;
            cellStyle.ShrinkToFit = shrinkToFit;

            return cellStyle;
        }

        private void CreateHomeSheet(Workbook excel, string sheetName, DBDto dbDto)
        {
            var sheet = excel.CreateSheet(sheetName);

            var titleStyle = CreateTitleStyle(excel);
            var cellStyle = CreateCellStyle(excel);

            var h1Style = CreateTitleStyle(excel, 24);
            var h2Style = CreateTitleStyle(excel, 18);

            sheet.CreateRange(0, 0, 2, 5).SetValue($"数据库文档  {dbDto.DBName}").SetStyle(h1Style);
            sheet.CreateRange(2, 0, 3, 5);


            int row = 5;
            sheet.CreateRange(row, 0, 1, 5).SetValue("修订日志").SetStyle(h2Style);

            row++;
            sheet.CreateCell(row, 0).SetValue("版本号").SetStyle(titleStyle);
            sheet.CreateCell(row, 1).SetValue("修订日期").SetStyle(titleStyle);
            sheet.CreateCell(row, 2).SetValue("修订内容").SetStyle(titleStyle);
            sheet.CreateCell(row, 3).SetValue("修订人").SetStyle(titleStyle);
            sheet.CreateCell(row, 4).SetValue("审核人").SetStyle(titleStyle);
            sheet.GetRow(row).Height = 28;

            for (var i = 0; i < 10; i++)
            {
                row++;
                sheet.CreateCell(row, 0).SetValue("").SetStyle(cellStyle);
                sheet.CreateCell(row, 1).SetValue("").SetStyle(cellStyle);
                sheet.CreateCell(row, 2).SetValue("").SetStyle(cellStyle);
                sheet.CreateCell(row, 3).SetValue("").SetStyle(cellStyle);
                sheet.CreateCell(row, 4).SetValue("").SetStyle(cellStyle);
                sheet.GetRow(row).Height = 20;
            }

            sheet.SetColumnWidth(0, 18);
            sheet.SetColumnWidth(1, 18);
            sheet.SetColumnWidth(2, 32);
            sheet.SetColumnWidth(3, 18);
            sheet.SetColumnWidth(4, 18);
        }

        private void CreateOverviewSheet(Workbook excel, string sheetName, DBDto dbDto)
        {
            var sheet = excel.CreateSheet(sheetName);

            var titleStyle = CreateTitleStyle(excel);
            var cellStyle = CreateCellStyle(excel);
            var cellLeftStyle = CreateCellStyle(excel, TextAlignmentType.Left);


            int row = 0;
            sheet.CreateCell(row, 0).SetValue("序号").SetStyle(titleStyle);
            sheet.CreateCell(row, 1).SetValue("表名").SetStyle(titleStyle);
            sheet.CreateCell(row, 2).SetValue("注释/说明").SetStyle(titleStyle);
            sheet.GetRow(row).Height = 28;

            for (var i = 0; i < dbDto.Tables.Count; i++)
            {
                row++;
                var table = dbDto.Tables[i];
                sheet.CreateCell(row, 0).SetValue(row).SetStyle(cellStyle);
                sheet.CreateCell(row, 1).SetValue(table.Name).SetStyle(cellLeftStyle);
                sheet.CreateCell(row, 2).SetValue(table.Comment??"").SetStyle(cellLeftStyle);
                sheet.GetRow(row).Height = 20;
            }

            sheet.SetColumnWidth(0, 10);
            sheet.SetColumnWidth(1, 50);
            sheet.SetColumnWidth(2, 60);
        }

        private void CreateTableSheet(Workbook excel, string sheetName, DBDto dbDto)
        {
            if (!dbDto.Tables.Any())
            {
                return;
            }

            var sheet = excel.CreateSheet(sheetName);

            var titleH1Style = CreateTitleStyle(excel, 16);
            var titleStyle = CreateTitleStyle(excel);
            var cellStyle = CreateCellStyle(excel);

            int row = 0;

            foreach (var table in dbDto.Tables)
            {
                sheet.CreateRange(row, 0, 1, 10).SetValue($"{table.Name} {table.Comment}").SetStyle(titleH1Style);
                sheet.GetRow(row).Height = 32;

                row++;
                sheet.CreateCell(row, 0).SetValue("序号").SetStyle(titleStyle);
                sheet.CreateCell(row, 1).SetValue("列名").SetStyle(titleStyle);
                sheet.CreateCell(row, 2).SetValue("数据类型").SetStyle(titleStyle);
                sheet.CreateCell(row, 3).SetValue("长度").SetStyle(titleStyle);
                sheet.CreateCell(row, 4).SetValue("小数位").SetStyle(titleStyle);
                sheet.CreateCell(row, 5).SetValue("主键").SetStyle(titleStyle);
                sheet.CreateCell(row, 6).SetValue("自增").SetStyle(titleStyle);
                sheet.CreateCell(row, 7).SetValue("允许空").SetStyle(titleStyle);
                sheet.CreateCell(row, 8).SetValue("默认值").SetStyle(titleStyle);
                sheet.CreateCell(row, 9).SetValue("列说明").SetStyle(titleStyle);
                sheet.GetRow(row).Height = 30;

                foreach (var column in table.Columns)
                {
                    row++;
                    sheet.CreateCell(row, 0).SetValue(column.ColumnOrder).SetStyle(cellStyle);
                    sheet.CreateCell(row, 1).SetValue(column.ColumnName).SetStyle(cellStyle);
                    sheet.CreateCell(row, 2).SetValue(column.ColumnType).SetStyle(cellStyle);
                    sheet.CreateCell(row, 3).SetValue(column.Length??"").SetStyle(cellStyle);
                    sheet.CreateCell(row, 4).SetValue(column.Scale??"").SetStyle(cellStyle);
                    sheet.CreateCell(row, 5).SetValue(column.IsPK ? "是" : "").SetStyle(cellStyle);
                    sheet.CreateCell(row, 6).SetValue(column.IsIdentity ? "是" : "").SetStyle(cellStyle);
                    sheet.CreateCell(row, 7).SetValue(column.CanNull ? "是" : "").SetStyle(cellStyle);
                    sheet.CreateCell(row, 8).SetValue(column.DefaultValue ?? "").SetStyle(cellStyle);
                    sheet.CreateCell(row, 9).SetValue(column.Comment ?? "").SetStyle(cellStyle);
                    sheet.GetRow(row).Height = 20;
                }

                row++;
                sheet.CreateRange(row, 0, 1, 10).SetValue("");
                sheet.GetRow(row).Height = 20;

                row++;
            }

            sheet.SetColumnWidth(0, 8);
            sheet.SetColumnWidth(1, 24);
            sheet.SetColumnWidth(2, 18);
            sheet.SetColumnWidth(3, 8);
            sheet.SetColumnWidth(4, 8);
            sheet.SetColumnWidth(5, 8);
            sheet.SetColumnWidth(6, 8);
            sheet.SetColumnWidth(7, 8);
            sheet.SetColumnWidth(8, 8);
            sheet.SetColumnWidth(9, 50);
        }

        private void CreateViewsSheet(Workbook excel, string sheetName, DBDto dbDto)
        {
            if (!dbDto.Views.Any())
            {
                return;
            }
            var sheet = excel.CreateSheet(sheetName);

            var titleStyle = CreateTitleStyle(excel);
            var cellLeftStyle = CreateCellStyle(excel, TextAlignmentType.Left, "Consolas", true);

            int row = 0;

            sheet.CreateCell(row, 0).SetValue("视图名").SetStyle(titleStyle);
            sheet.CreateCell(row, 1).SetValue("视图定义").SetStyle(titleStyle);
            sheet.GetRow(row).Height = 30;

            foreach (var view in dbDto.Views)
            {
                row++;
                sheet.CreateCell(row, 0).SetValue(view.Name).SetStyle(cellLeftStyle);
                sheet.CreateCell(row, 1).SetValue(view.Definition).SetStyle(cellLeftStyle);
            }

            sheet.SetColumnWidth(0, 24);
            sheet.SetColumnWidth(1, 50);
        }

        private void CreateProcsSheet(Workbook excel, string sheetName, DBDto dbDto)
        {
            if (!dbDto.Procs.Any())
            {
                return;
            }
            var sheet = excel.CreateSheet(sheetName);

            var titleStyle = CreateTitleStyle(excel);
            var cellLeftStyle = CreateCellStyle(excel, TextAlignmentType.Left, "Consolas", true);

            int row = 0;

            sheet.CreateCell(row, 0).SetValue("存储过程名").SetStyle(titleStyle);
            sheet.CreateCell(row, 1).SetValue("存储过程定义").SetStyle(titleStyle);
            sheet.GetRow(row).Height = 30;

            foreach (var proc in dbDto.Procs)
            {
                row++;
                sheet.CreateCell(row, 0).SetValue(proc.Name).SetStyle(cellLeftStyle);
                sheet.CreateCell(row, 1).SetValue(proc.Definition).SetStyle(cellLeftStyle);
            }

            sheet.SetColumnWidth(0, 24);
            sheet.SetColumnWidth(1, 50);
        }
    }
}