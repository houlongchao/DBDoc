using MySqlConnector;
using System.Data.Common;

namespace DBDoc.DB.MySql
{
    public class MySqlDB : BaseDB
    {
        public new const string DBType = "MySql";

        private string dbVersionSql = "select @@version";
        private string dbNamesSql = "SELECT SCHEMA_NAME as name FROM information_schema.SCHEMATA where SCHEMA_NAME not in ('information_schema', 'performance_schema', 'mysql', 'sys') order by SCHEMA_NAME asc";
        private string schemasSql = "";
        private string tablesSql = "SELECT table_name as name,TABLE_COMMENT as comment FROM INFORMATION_SCHEMA.TABLES WHERE lower(table_type)='base table' and  table_schema =@dbName order by table_name asc ";
        private string viewsSql = "SELECT  table_name as name,VIEW_DEFINITION as definition FROM  information_schema.views where TABLE_SCHEMA=@dbName order by table_name asc";
        private string procsSql = "SELECT  ROUTINE_NAME as name,ROUTINE_DEFINITION as definition FROM INFORMATION_SCHEMA.ROUTINES WHERE ROUTINE_SCHEMA=@dbName AND ROUTINE_TYPE='PROCEDURE' ORDER BY ROUTINE_NAME ASC";
        private string columnSql = @"select ORDINAL_POSITION as ColumnOrder,Column_Name as ColumnName,data_type as ColumnType,COLUMN_COMMENT as Comment,
(case when data_type = 'float' or data_type = 'double' or data_type = 'decimal' then  NUMERIC_PRECISION else CHARACTER_MAXIMUM_LENGTH end ) as length,
(case when data_type = 'int' or data_type = 'bigint' then null else NUMERIC_SCALE end) as Scale,( case when EXTRA='auto_increment' then 1 else 0 end) as IsIdentity,(case when COLUMN_KEY='PRI' then 1 else 0 end) as IsPK,
(case when IS_NULLABLE = 'NO' then 0 else 1 end) as CanNull,COLUMN_DEFAULT as DefaultValue
from information_schema.columns where table_schema = @dbName and table_name = @tableName order by ORDINAL_POSITION asc";

        public MySqlDB(string connectionString, int timeOut = 3000) : base(DBType, connectionString)
        {
            this.Timeout = timeOut;
        }

        public override string DBName
        {
            get { return new MySqlConnectionStringBuilder(ConnectionString).Database; }
        }

        public override DbConnection CreateConnection()
        {
            return new MySqlConnection(ConnectionString);
        }

        public override string GetColumnsSql()
        {
            return columnSql;
        }

        public override string GetDbNamesSql()
        {
            return dbNamesSql;
        }

        public override string GetDbVersionSql()
        {
            return dbVersionSql;
        }

        public override string GetProcsSql()
        {
            return procsSql;
        }

        public override string GetSchemasSql()
        {
            return schemasSql;
        }

        public override string GetTablesSql()
        {
            return tablesSql;
        }

        public override string GetViewsSql()
        {
            return viewsSql;
        }
    }
}