using Aspose.Words;
using Aspose.Words.Tables;
using DBDoc.Dto;
namespace DBDoc.Doc.Word
{
    public partial class WordDoc
    {
        private static int FontSize_Level_1 = 32;
        private static int FontSize_Level_2 = 24;
        private static int FontSize_Level_3 = 16;

        public static void ExportWordByAsposeWords(string fileName, DBDto dbDto)
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

            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            CreateHome(builder, "修订日志", dbDto);
            CreateOverview(builder, "数据库表目录", dbDto);
            CreateTables(builder, "数据库表结构", dbDto);
            CreateViews(builder, "数据库视图", dbDto);
            CreateProcs(builder, "数据库存储过程", dbDto);
            
            AutoGenPageNum(doc, builder);

            doc.Save(fileName);
        }

        private static void CreateHome(DocumentBuilder builder, string title, DBDto dbDto)
        {
            WriteEmptyLine(builder, 5);
            WriteTitle(builder, "数据库文档", OutlineLevel.Level1, FontSize_Level_1, ParagraphAlignment.Center);
            WriteText(builder, dbDto.DBName, 22);

            WriteEmptyLine(builder, 3);
            WriteTitle(builder, title, OutlineLevel.Level2, FontSize_Level_2, ParagraphAlignment.Center);

            Table logTable = builder.StartTable();

            SetTableTitleStyle(builder);

            builder.InsertCell();
            builder.Write("版本号");

            builder.InsertCell();
            builder.Write("修订日期");

            builder.InsertCell();
            builder.Write("修订内容");

            builder.InsertCell();
            builder.Write("修订人");

            builder.InsertCell();
            builder.Write("审核人");

            builder.EndRow();

            for (var i = 0; i < 5; i++)
            {
                SetTableDataStyle(builder);

                builder.InsertCell();
                builder.Write(""); // 版本号

                builder.InsertCell();
                builder.Write(""); // 修订日期

                builder.InsertCell();
                builder.Write(""); // 修订内容

                builder.InsertCell();
                builder.Write(""); // 修订人

                builder.InsertCell();
                builder.Write(""); // 审核人

                builder.EndRow();
            }


            logTable.Alignment = TableAlignment.Center;
            logTable.AllowAutoFit = true;

            builder.EndTable();
        }

        private static void CreateOverview(DocumentBuilder builder, string title, DBDto dbDto)
        {
            WriteNewPage(builder);
            WriteTitle(builder, title, OutlineLevel.Level2, FontSize_Level_2, ParagraphAlignment.Center);

            Table overviewTable = builder.StartTable();

            SetTableTitleStyle(builder);

            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPercent(50);
            builder.Write("表名");

            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPercent(50);
            builder.Write("注释/说明");

            builder.EndRow();

            foreach (var table in dbDto.Tables)
            {
                SetTableDataStyle(builder, ParagraphAlignment.Left, "Consolas");

                builder.InsertCell();
                builder.Write(table.Name); // 表名

                builder.InsertCell();
                builder.Write((!string.IsNullOrWhiteSpace(table.Comment) ? table.Comment : "")); // 说明


                builder.EndRow();
            }

            SetTableStyle(overviewTable);

            builder.EndTable();
        }

        private static void CreateTables(DocumentBuilder builder, string title, DBDto dbDto)
        {
            WriteNewPage(builder);
            WriteTitle(builder, title, OutlineLevel.Level2, FontSize_Level_2, ParagraphAlignment.Center);

            foreach (var table in dbDto.Tables)
            {
                string tableTitle = $"{table.Name} {table.Comment ?? ""}";

                WriteTitle(builder, tableTitle, OutlineLevel.Level3, FontSize_Level_3);

                Table asposeTable = builder.StartTable();

                builder.ParagraphFormat.ClearFormatting();

                SetTableTitleStyle(builder);

                builder.InsertCell();
                builder.CellFormat.PreferredWidth = PreferredWidth.FromPercent(6);
                builder.Write("序号");

                builder.InsertCell();
                builder.CellFormat.PreferredWidth = PreferredWidth.FromPercent(20);
                builder.Write("列名");

                builder.InsertCell();
                builder.CellFormat.PreferredWidth = PreferredWidth.FromPercent(15);
                builder.Write("数据类型");

                builder.InsertCell();
                builder.CellFormat.PreferredWidth = PreferredWidth.FromPercent(10);
                builder.Write("长度");

                builder.InsertCell();
                builder.CellFormat.PreferredWidth = PreferredWidth.FromPercent(6);
                builder.Write("小数位");

                builder.InsertCell();
                builder.CellFormat.PreferredWidth = PreferredWidth.FromPercent(6);
                builder.Write("主键");

                builder.InsertCell();
                builder.CellFormat.PreferredWidth = PreferredWidth.FromPercent(6);
                builder.Write("自增");

                builder.InsertCell();
                builder.CellFormat.PreferredWidth = PreferredWidth.FromPercent(6);
                builder.Write("允许空");

                builder.InsertCell();
                builder.CellFormat.PreferredWidth = PreferredWidth.FromPercent(10);
                builder.Write("默认值");

                builder.InsertCell();
                builder.CellFormat.PreferredWidth = PreferredWidth.FromPercent(30);
                builder.Write("列说明");
                builder.EndRow();

                foreach (var column in table.Columns)
                {
                    SetTableDataStyle(builder);

                    builder.InsertCell();
                    builder.Write(column.ColumnOrder.ToString()); // 序号

                    builder.InsertCell();
                    builder.Write(column.ColumnName); // 列名

                    builder.InsertCell();
                    builder.Write(column.ColumnType); // 数据类型

                    builder.InsertCell();
                    builder.Write(column.Length ?? ""); // 长度

                    builder.InsertCell();
                    builder.Write(column.Scale ?? ""); // 小数位

                    builder.InsertCell();
                    builder.Write(column.IsPK ? "是" : ""); // 主键

                    builder.InsertCell();
                    builder.Write(column.IsIdentity ? "是" : ""); // 自增

                    builder.InsertCell();
                    builder.Write(column.CanNull ? "是" : ""); // 是否为空

                    builder.InsertCell();
                    builder.Write(column.DefaultValue ?? ""); // 默认值

                    builder.InsertCell();
                    builder.Write(column.Comment ?? ""); // 列说明

                    builder.EndRow();
                }

                SetTableStyle(asposeTable);

                builder.EndTable();

                WriteNewPage(builder);
            }
        }

        private static void CreateViews(DocumentBuilder builder, string title, DBDto dbDto)
        {
            if (!dbDto.Views.Any())
            {
                return;
            }

            WriteNewPage(builder);
            WriteTitle(builder, title, OutlineLevel.Level2, FontSize_Level_2, ParagraphAlignment.Center);

            Table overviewTable = builder.StartTable();

            SetTableTitleStyle(builder);

            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPercent(30);
            builder.Write("视图名");

            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPercent(70);
            builder.Write("视图定义");

            builder.EndRow();

            foreach (var view in dbDto.Views)
            {
                SetTableDataStyle(builder, ParagraphAlignment.Left, "Consolas");

                builder.InsertCell();
                builder.Write(view.Name); 

                builder.InsertCell();
                builder.Write(view.Definition ?? "");


                builder.EndRow();
            }

            SetTableStyle(overviewTable);

            builder.EndTable();
        }

        private static void CreateProcs(DocumentBuilder builder, string title, DBDto dbDto)
        {
            if (!dbDto.Procs.Any())
            {
                return;
            }

            WriteNewPage(builder);
            WriteTitle(builder, title, OutlineLevel.Level2, FontSize_Level_2, ParagraphAlignment.Center);

            Table overviewTable = builder.StartTable();

            
            SetTableTitleStyle(builder);

            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPercent(30);
            builder.Write("存储过程名");

            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPercent(70);
            builder.Write("存储过程定义");

            builder.EndRow();

            foreach (var proc in dbDto.Procs)
            {
                SetTableDataStyle(builder, ParagraphAlignment.Left, "Consolas");

                builder.InsertCell();
                builder.Write(proc.Name);

                builder.InsertCell();
                builder.Write(proc.Definition ?? "");


                builder.EndRow();
            }

            SetTableStyle(overviewTable);

            builder.EndTable();
        }

        /// <summary>
        /// 生成页码
        /// </summary>
        /// <param name="builder"></param>
        public static void AutoGenPageNum(Document doc, DocumentBuilder builder)
        {
            HeaderFooter footer = new HeaderFooter(doc, HeaderFooterType.FooterPrimary);
            doc.FirstSection.HeadersFooters.Add(footer);

            footer.AppendParagraph("").ParagraphFormat.Alignment = ParagraphAlignment.Right;
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
            builder.InsertField("PAGE");
            builder.Write(" / ");
            builder.InsertField("NUMPAGES");
        }

        private static void SetTableStyle(Table table)
        {
            table.Alignment = TableAlignment.Center;
            table.PreferredWidth = PreferredWidth.FromPercent(100);
            table.AllowAutoFit = false;
        }

        private static void SetTableTitleStyle(DocumentBuilder builder)
        {
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;

            builder.RowFormat.Height = 32;
            builder.RowFormat.HeightRule = HeightRule.AtLeast;

            builder.CellFormat.Shading.BackgroundPatternColor = System.Drawing.Color.FromArgb(198, 217, 241);
            builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;

            builder.Font.Size = 12;
            builder.Font.Name = "Arial";
            builder.Font.Bold = true;
        }

        private static void SetTableDataStyle(DocumentBuilder builder, ParagraphAlignment alignment = ParagraphAlignment.Center, string fontName = "Arial")
        {
            builder.ParagraphFormat.Alignment = alignment;

            builder.RowFormat.Height = 24.0;
            builder.RowFormat.HeightRule = HeightRule.AtLeast;

            builder.CellFormat.Shading.BackgroundPatternColor = System.Drawing.Color.White;
            builder.CellFormat.Width = 50.0;
            builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;

            builder.Font.Size = 10;
            builder.Font.Name = fontName;
            builder.Font.Bold = false;
        }

        /// <summary>
        /// 写标题
        /// </summary>
        /// <param name="builder"></param>
        /// <param name="title"></param>
        /// <param name="outlineLevel"></param>
        /// <param name="fontSize"></param>
        /// <param name="alignment"></param>
        private static void WriteTitle(DocumentBuilder builder, string title, OutlineLevel outlineLevel, double fontSize, ParagraphAlignment alignment = ParagraphAlignment.Left)
        {
            builder.ParagraphFormat.OutlineLevel = outlineLevel;
            builder.ParagraphFormat.Alignment = alignment;
            builder.Font.Size = fontSize;
            builder.Font.Name = "Arial";
            builder.Font.Bold = true;
            builder.Writeln(title);
        }

        /// <summary>
        /// 写内容
        /// </summary>
        /// <param name="builder"></param>
        /// <param name="text"></param>
        /// <param name="fontSize"></param>
        private static void WriteText(DocumentBuilder builder, string text, double fontSize)
        {
            builder.ParagraphFormat.OutlineLevel = OutlineLevel.BodyText;
            builder.Font.Size = fontSize;
            builder.Font.Name = "Arial";
            builder.Writeln(text);
        }

        /// <summary>
        /// 写空行
        /// </summary>
        /// <param name="builder"></param>
        /// <param name="line"></param>
        private static void WriteEmptyLine(DocumentBuilder builder, int line = 1)
        {
            builder.ParagraphFormat.OutlineLevel = OutlineLevel.BodyText;
            for (int i = 0; i < line; i++)
            {
                builder.InsertBreak(BreakType.ParagraphBreak);
            }
        }

        /// <summary>
        /// 写新页
        /// </summary>
        /// <param name="builder"></param>
        private static void WriteNewPage(DocumentBuilder builder)
        {
            builder.ParagraphFormat.OutlineLevel = OutlineLevel.BodyText;
            builder.InsertBreak(BreakType.PageBreak);
        }
    }
}
