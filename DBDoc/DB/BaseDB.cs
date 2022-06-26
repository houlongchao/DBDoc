using DBDoc.Dto;
using System.Data.Common;
using Dapper;

namespace DBDoc.DB
{
    public abstract class BaseDB : IDB
    {
        public BaseDB(string dbType, string connectionString)
        {
            this.DBType = dbType;
            this.ConnectionString = connectionString;
        }

        /// <summary>
        /// 数据库类型
        /// </summary>
        public string DBType { get; set;}

        /// <summary>
        /// 数据库名称
        /// </summary>
        public abstract string DBName { get; }


        /// <summary>
        /// 超时时间
        /// </summary>
        public int Timeout { get; set; }

        /// <summary>
        /// 连接字符串
        /// </summary>
        public string ConnectionString { get; set; }

        /// <summary>
        /// 数据库连接
        /// </summary>
        public DbConnection DbConnection => CreateConnection();

        /// <summary>
        /// 创建 Connection 对象
        /// </summary>
        /// <returns></returns>
        public abstract DbConnection CreateConnection();

        public abstract string GetColumnsSql();

        public virtual List<ColumnDto> GetColumns(string tableName)
        {
            return DbConnection.Query<ColumnDto>(GetColumnsSql(), new { dbName = DBName, tableName = tableName }).ToList();
        }

        public abstract string GetDbNamesSql();
        public virtual List<string> GetDbNames()
        {
            return DbConnection.Query<string>(GetDbNamesSql()).ToList();
        }

        public abstract string GetDbVersionSql();
        public virtual string GetDbVersion()
        {
            return DbConnection.QueryFirst<string>(GetDbVersionSql());
        }

        public abstract string GetProcsSql();
        public virtual List<ProcDto> GetProcs()
        {
            return DbConnection.Query<ProcDto>(GetProcsSql(), new { dbName = DBName }).ToList();
        }

        public abstract string GetSchemasSql();
        public virtual List<string> GetSchemas()
        {
            return DbConnection.Query<string>(GetSchemasSql(), new { dbName = DBName }).ToList();
        }

        public abstract string GetTablesSql();
        public virtual List<TableDto> GetTables()
        {
            var a = GetTablesSql();
            return DbConnection.Query<TableDto>(GetTablesSql(), new { dbName = DBName }).ToList();
        }

        public abstract string GetViewsSql();
        public virtual List<ViewDto> GetViews()
        {
            return DbConnection.Query<ViewDto>(GetViewsSql(), new { dbName = DBName }).ToList();
        }
    }
}
