namespace DBDoc.Dto
{
    public class TableDto
    {
        /// <summary>
        /// 表名
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 注释
        /// </summary>
        public string? Comment { get; set; }

        /// <summary>
        /// Schema
        /// </summary>
        public string? Schema { get; set; }

        /// <summary>
        /// 表格列数据
        /// </summary>
        public List<ColumnDto> Columns { get; set; } = new List<ColumnDto>();

    }
}
