using DBDoc.Dto;

namespace DBDoc.Doc.Excel
{
    public partial class ExcelDoc : BaseDoc
    {
        public const string DocType = "Excel";

        public ExcelDoc(DBDto dbDto) : base(dbDto)
        {
        }

        public override void Build(string filePath)
        {
            ExportExcelByAspose(filePath, DbDto);
        }
    }
}