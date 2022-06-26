namespace DBDoc.Dto
{
    public class DBDto
    {
        public DBDto(string dbName)
        {
            this.DBName = dbName;
        }

        /// <summary>
        /// 数据库名称
        /// </summary>
        public string DBName { get; set; }

        /// <summary>
        /// 数据库类型
        /// </summary>
        public string DBType { get; set; }

        /// <summary>
        /// 表结构信息
        /// </summary>
        public List<TableDto> Tables { get; set; } = new List<TableDto>();

        /// <summary>
        /// 数据库视图
        /// </summary>
        public List<ViewDto> Views { get; set; } = new List<ViewDto>();

        /// <summary>
        /// 数据库存储过程
        /// </summary>
        public List<ProcDto> Procs { get; set; } = new List<ProcDto>();
    }
}
