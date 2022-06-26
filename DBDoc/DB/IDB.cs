using DBDoc.Dto;

namespace DBDoc.DB
{
    public interface IDB
    {
        /// <summary>
        /// 获取DB版本
        /// </summary>
        /// <returns></returns>
        string GetDbVersion();

        /// <summary>
        /// 获取DB
        /// </summary>
        /// <returns></returns>
        List<string> GetDbNames();

        /// <summary>
        /// 获取Schema
        /// </summary>
        /// <returns></returns>
        public List<string> GetSchemas();

        /// <summary>
        /// 获取表
        /// </summary>
        /// <returns></returns>
        List<TableDto> GetTables();

        /// <summary>
        /// 获取视图
        /// </summary>
        /// <returns></returns>
        List<ViewDto> GetViews();

        /// <summary>
        /// 获取存储过程
        /// </summary>
        /// <returns></returns>
        List<ProcDto> GetProcs();

        /// <summary>
        /// 获取列
        /// </summary>
        /// <param name="tableName"></param>
        /// <returns></returns>
        List<ColumnDto> GetColumns(string tableName);
    }
}
