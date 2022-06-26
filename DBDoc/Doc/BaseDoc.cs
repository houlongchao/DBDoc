using DBDoc.Dto;

namespace DBDoc.Doc
{
    /// <summary>
    /// 文档生成基类
    /// </summary>
    public abstract class BaseDoc
    {
        public BaseDoc(DBDto dbDto)
        {
            this.DbDto = dbDto;
        }

        /// <summary>
        /// 文件扩展名
        /// </summary>
        public string Ext { get; set; }

        /// <summary>
        /// 数据库Dto
        /// </summary>
        public DBDto DbDto { get; }

        /// <summary>
        /// 构建生成文档
        /// </summary>
        public abstract void Build(string filePath);
    }
}
