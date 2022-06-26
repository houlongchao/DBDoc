using Npgsql;
using System.Data.Common;

namespace DBDoc.DB.Postgres
{
    public class PostgresDB : BaseDB
    {
        public new const string DBType = "Postgres";

        private string dbVersionSql = "select version()";
        private string dbNamesSql = "select datname as name from pg_database where datistemplate=false order by oid desc";
        private string schemasSql = "";
        private string tablesSql = @"select a.*,cast(obj_description(relfilenode,'pg_class') as varchar) as comment from (
select b.oid as schid, a.table_schema as schema,concat(a.table_schema,'.',a.table_name) as Name,a.table_name from information_schema.tables  a
left join pg_namespace b on b.nspname = a.table_schema
where a.table_schema not in ('pg_catalog','information_schema') and a.table_type='BASE TABLE'
) a inner join pg_class b on a.table_name = b.relname and a.schid = b.relnamespace
order by name asc";
        private string viewsSql = "SELECT viewname as name,definition FROM pg_views where schemaname not in ('pg_catalog','information_schema') order by viewname asc";
        private string procsSql = "select proname as name,prosrc as definition from  pg_proc where pronamespace in (SELECT pg_namespace.oid FROM pg_namespace WHERE nspname not in ('pg_catalog','information_schema') ) order by proname asc";
        private string columnSql = @"SELECT
	col.ordinal_position AS ColumnOrder,
	concat(col.table_schema,'.', col.TABLE_NAME) as tableName,
	col.COLUMN_NAME AS columnName,
	col.udt_name AS ColumnType,
	col_description ( C.oid, col.ordinal_position ) AS Comment,	
	COALESCE(col.character_maximum_length, col.numeric_precision, -1) AS Length,
	col.numeric_scale AS Scale,
	(CASE col.is_identity WHEN 'YES' THEN 1 ELSE 0 END) AS IsIdentity,	
	(CASE col.is_nullable WHEN 'YES' THEN 1 ELSE 0 END) AS CanNull,
	col.column_default AS DefaultVal
FROM
	information_schema.COLUMNS AS col
	LEFT JOIN pg_namespace ns ON ns.nspname = col.table_schema
	LEFT JOIN pg_class C ON col.TABLE_NAME = C.relname 
	AND C.relnamespace = ns.oid
	LEFT JOIN pg_attrdef A ON C.oid = A.adrelid 
	AND col.ordinal_position = A.adnum
	LEFT JOIN pg_attribute b ON b.attrelid = C.oid 
	AND b.attname = col.
	COLUMN_NAME LEFT JOIN pg_type et ON et.oid = b.atttypid
	LEFT JOIN pg_collation coll ON coll.oid = b.attcollation
	LEFT JOIN pg_namespace colnsp ON coll.collnamespace = colnsp.oid
	LEFT JOIN pg_type bt ON et.typelem = bt.oid
	LEFT JOIN pg_namespace nbt ON bt.typnamespace = nbt.oid
WHERE
    concat(col.table_schema,'.', col.TABLE_NAME) = @tableName
ORDER BY
	tableName,
	col.ordinal_position";

        public PostgresDB(string connectionString, int timeOut = 3000) : base(DBType, connectionString)
        {
            this.Timeout = timeOut;
        }

        public override string DBName
        {
            get { return new NpgsqlConnectionStringBuilder(ConnectionString).Database ?? ""; }
        }

        public override DbConnection CreateConnection()
        {
            return new NpgsqlConnection(ConnectionString);
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