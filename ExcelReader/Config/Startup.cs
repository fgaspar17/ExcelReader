using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReader;

public class Startup
{
    public static void StartApplication()
    {
        var builder = new ConfigurationBuilder()
            .SetBasePath(AppContext.BaseDirectory)
            .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);

        IConfiguration config = builder.Build();

        string connectionString = config.GetConnectionString("Sqlite") 
            ?? throw new InvalidOperationException("You must provide a Sqlite Connection String in the appsettings.json file.");
        GlobalConfig.InitializeConnectionString(connectionString);

        string dbFile = config.GetValue<string>("DatabaseFileName")
            ?? throw new InvalidOperationException("You must provide a DatabaseFileName in the appsettings.json file.");

        SetupDatabase.ResetDatabase(dbFile);
    }
}
