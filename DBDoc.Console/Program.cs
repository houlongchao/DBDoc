using DBDoc.DB;
using DBDoc.DB.MySql;
using DBDoc.DB.Postgres;
using DBDoc.DB.SqlServer;
using DBDoc.Doc.Excel;
using DBDoc.Doc.Word;
using DBDoc.Dto;
using Spectre.Console;
using System.Diagnostics;

try
{
    AnsiConsole.Write(new FigletText("Welcome").Color(Color.Blue));
    AnsiConsole.Write(new FigletText("      DBDoc").Color(Color.Blue));

    // DB Type
    var dbType = AnsiConsole.Prompt(new SelectionPrompt<string>().Title("[green]DB Type[/]:").AddChoices(new[] { MySqlDB.DBType, PostgresDB.DBType, SqlServerDB.DBType }));
    AnsiConsole.Markup($"[green]DB Type[/]: {dbType}{Environment.NewLine}");

    // DB Host
    var dbHost = AnsiConsole.Ask<string>("[green]DB Host[/]:");

    // DB Port
    var dbPort = AnsiConsole.Ask<string>("[green]DB Port[/]:");

    // DB User
    var dbUser = AnsiConsole.Ask<string>("[green]DB User[/]:");

    // DB Pass
    var dbPass = AnsiConsole.Prompt(new TextPrompt<string>("[green]DB Password[/]:").Secret());

    // Get DB Names
    BaseDB? db = default;
    switch (dbType)
    {
        case MySqlDB.DBType:
            {
                db = new MySqlDB($"Data Source={dbHost};Port={dbPort};User ID={dbUser};Password={dbPass};");
                break;
            }
        case PostgresDB.DBType:
            {
                db = new PostgresDB($"Host={dbHost};Port={dbPort};Username={dbUser};Password={dbPass};");
                break;
            }
        case SqlServerDB.DBType:
            {
                db = new SqlServerDB($"Data Source={dbHost},{dbPort};User Id={dbUser};Password={dbPass};");
                break;
            }
    }
    var dbNames = new List<string>();
    AnsiConsole.Status().Start("Running...", ctx =>
    {
        ctx.Status("Get DB Names...");
        dbNames = db?.GetDbNames();
    });

    // DB Name
    var dbName = AnsiConsole.Prompt(new SelectionPrompt<string>().Title("[green]DB Name[/]:").AddChoices(dbNames));
    AnsiConsole.Markup($"[green]DB Name[/]: {dbName}{Environment.NewLine}");

    // Output Types
    var outputTypes = AnsiConsole.Prompt(new MultiSelectionPrompt<string>()
            .Title("[green]Output Type[/]:")
            .InstructionsText("[grey](Press [blue]<space>[/] to toggle, [green]<enter>[/] to accept)[/]")
            .Select("Word")
            .AddChoices(new[] { WordDoc.DocType, ExcelDoc.DocType }));
    foreach (string outputType in outputTypes)
    {
        AnsiConsole.Markup($"[green]Output Type[/]: {outputType}{Environment.NewLine}");
    }

    // Get DB Info & Output
    switch (dbType)
    {
        case MySqlDB.DBType:
            {
                db = new MySqlDB($"Data Source={dbHost};Port={dbPort};User ID={dbUser};Password={dbPass}; Initial Catalog={dbName};Charset=utf8; SslMode=none;");
                break;
            }
        case PostgresDB.DBType:
            {
                db = new PostgresDB($"Host={dbHost};Port={dbPort};Username={dbUser};Password={dbPass}; Database={dbName};");
                break;
            }
        case SqlServerDB.DBType:
            {
                db = new SqlServerDB($"Data Source={dbHost},{dbPort};User Id={dbUser};Password={dbPass};Initial Catalog={dbName};TrustServerCertificate=true;");
                break;
            }
    }
    AnsiConsole.Status().Start("Running...", ctx =>
        {
            ctx.Status("Get DB Info...");
            var dbDto = new DBDto(db.DBName);
            dbDto.Tables = db.GetTables();
            dbDto.Views = db.GetViews();
            dbDto.Procs = db.GetProcs();
            foreach (var table in dbDto.Tables)
            {
                table.Columns = db.GetColumns(table.Name);
            }

            var myDocumentsFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            var fileFolder = Path.Combine(myDocumentsFolder, "DBDoc");
            Directory.CreateDirectory(fileFolder);

            var dateTime = DateTime.Now.ToString("yyMMddHHmm");
            foreach (string docType in outputTypes)
            {
                switch (docType)
                {
                    case WordDoc.DocType:
                        {
                            ctx.Status("Write To Word...");
                            ctx.Spinner(Spinner.Known.Star);
                            var filePath = Path.Combine(fileFolder, $"{dbDto.DBName}_{dateTime}.docx");
                            var doc = new WordDoc(dbDto);
                            doc.Build(filePath);
                            AnsiConsole.Markup($"[green]Output[/]: {filePath}{Environment.NewLine}");
                            break;
                        }
                    case ExcelDoc.DocType:
                        {
                            ctx.Status("Write To Excel...");
                            ctx.Spinner(Spinner.Known.Star);
                            var filePath = Path.Combine(fileFolder, $"{dbDto.DBName}_{dateTime}.xlsx");
                            var excel = new ExcelDoc(dbDto);
                            excel.Build(filePath);
                            AnsiConsole.Markup($"[green]Output[/]: {filePath}{Environment.NewLine}");
                            break;
                        }
                }
            }

            Process.Start("explorer.exe", fileFolder);
        });

    AnsiConsole.Markup($"[blue]Finished!!![/] [grey](Enter any key to end)[/]");
    Console.ReadKey();
}
catch (Exception ex)
{
    AnsiConsole.WriteException(ex);
    Console.ReadKey();
}