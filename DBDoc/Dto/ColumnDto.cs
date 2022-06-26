namespace DBDoc.Dto
{
    public class ColumnDto
    {
        /// <summary>
        /// 列序号
        /// </summary>
        public int ColumnOrder { get; set; }

        /// <summary>
        /// 列名
        /// </summary>
        public string ColumnName { get; set; }

        /// <summary>
        /// 数据类型
        /// </summary>
        public string ColumnType { get; set; }

        /// <summary>
        /// 长度
        /// </summary>
        public string? Length { get; set; }

        /// <summary>
        /// 小数位
        /// </summary>
        public string? Scale { get; set; }

        /// <summary>
        /// 主键
        /// </summary>
        public bool IsPK { get; set; }

        /// <summary>
        /// 自增
        /// </summary>
        public bool IsIdentity { get; set; }

        /// <summary>
        /// 允许空
        /// </summary>
        public bool CanNull { get; set; }

        /// <summary>
        /// 默认值
        /// </summary>
        public string? DefaultValue { get; set; }

        /// <summary>
        /// 注释
        /// </summary>
        public string? Comment { get; set; }

    }
}
