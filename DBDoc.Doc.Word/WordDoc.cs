using DBDoc.Dto;

namespace DBDoc.Doc.Word
{
    public partial class WordDoc : BaseDoc
    {
        public const string DocType = "Word";

        public WordDoc(DBDto dbDto) : base(dbDto)
        {
        }

        public override void Build(string filePath)
        {
            ExportWordByAsposeWords(filePath, DbDto);
        }
    }
}