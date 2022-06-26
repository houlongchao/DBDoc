using System.Data.Common;
using System.Data.SqlClient;

namespace DBDoc.DB.SqlServer
{
    public class SqlServerDB : BaseDB
    {
        public new const string DBType = "SqlServer";

        private string dbVersionSql = "select @@version";
        private string dbNamesSql = "select name from sys.sysdatabases where name not in ('master', 'model', 'msdb', 'tempdb') Order By name asc";
        private string schemasSql = "Select distinct a.name From sys.schemas a left join sys.objects b on a.schema_id = b.schema_id WHERE b.type = 'U'";
        private string tablesSql = "SELECT * FROM (SELECT (Select Top 1 c.name From sys.schemas c Where c.schema_id = a.schema_id) + '.' + a.Name + '' as Name,(SELECT TOP 1 Value FROM sys.extended_properties b WHERE b.major_id=a.object_id and b.minor_id=0) AS comment,(Select top 1 c.name From sys.schemas c Where c.schema_id = a.schema_id) as [schema] From sys.objects a WHERE a.type = 'U' AND a.name <> 'sysdiagrams' AND a.name <> 'dtproperties' )K WHERE K.[schema] <> 'cdc' ORDER BY K.name ASC;";
        private string viewsSql = "SELECT TABLE_NAME as name,VIEW_DEFINITION as definition FROM INFORMATION_SCHEMA.VIEWS Order By TABLE_NAME asc";
        private string procsSql = "select name,[definition] from sys.objects a Left Join sys.sql_modules b On a.[object_id]=b.[object_id] Where a.type='P' And a.is_ms_shipped=0 And b.execute_as_principal_id Is Null And name !='sp_upgraddiagrams' Order By a.name asc";
        private string columnSql = @"select s.* FROM (SELECT a.colorder ColumnOrder,a.name ColumnName,b.name ColumnType,row_number() over (partition by a.name order by b.name) as group_idx,(case when (SELECT count(*) FROM sysobjects  WHERE (name in (SELECT name FROM sysindexes  WHERE (id = a.id) AND (indid in  (SELECT indid FROM sysindexkeys  WHERE (id = a.id) AND (colid in  (SELECT colid FROM syscolumns WHERE (id = a.id) AND (name = a.name)))))))  AND (xtype = 'PK'))>0 then 1 else 0 end) IsPK,(case when COLUMNPROPERTY( a.id,a.name,'IsIdentity')=1 then 1 else 0 end) IsIdentity,  CASE When b.name ='uniqueidentifier' Then 36  WHEN (charindex('int',b.name)>0) OR (charindex('time',b.name)>0) THEN NULL ELSE  COLUMNPROPERTY(a.id,a.name,'PRECISION') end as [Length], CASE WHEN ((charindex('int',b.name)>0) OR (charindex('time',b.name)>0)) THEN NULL ELSE isnull(COLUMNPROPERTY(a.id,a.name,'Scale'),null) end as Scale,(case when a.isnullable=1 then 1 else 0 end) CanNull,Replace(Replace(IsNull(e.text,''),'(',''),')','') DefaultValue,isnull(g.[value], ' ') AS Comment FROM  syscolumns a left join systypes b on a.xtype=b.xusertype inner join sysobjects d on a.id=d.id and d.xtype='U' and d.name<>'dtproperties' left join syscomments e on a.cdefault=e.id left join sys.extended_properties g on a.id=g.major_id AND a.colid=g.minor_id  And g.class=1 left join sys.extended_properties f on d.id=f.class and f.minor_id=0 left join sys.schemas c on d.uid = c.schema_id where b.name is not NULL and  c.name + '.' + d.name = @tableName ) s WHERE s.group_idx = 1 order by s.ColumnOrder";

        public SqlServerDB(string connectionString, int timeOut = 3000) : base(DBType, connectionString)
        {
            this.Timeout = timeOut;
        }

        public override string DBName
        {
            get { return new SqlConnectionStringBuilder(ConnectionString).InitialCatalog; }
        }

        public override DbConnection CreateConnection()
        {
            return new SqlConnection(ConnectionString);
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