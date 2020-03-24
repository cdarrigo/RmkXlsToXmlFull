using System;
using System.IO;
using CommandLine;
using Newtonsoft.Json;
using Serilog;
using Serilog.Events;

namespace RmkXlsToXml
{
    class Program
    {
        static void Main(string[] args)
        {
           
            ConfigureLogger();
            // parse the command line args into a strongly typed ConverterConfiguration instance
            Parser.Default.ParseArguments<ConverterConfiguration>(args)
                .WithParsed(config =>
                {
                    try
                    {
                        if (!config.Validate())
                        {
                            HandleError("Invalid configuration.");
                            return;
                        }

                        Log.Logger.Information($"Converting data from {config.SourceFile} to xml.");
                        var converter = new RemarketingDataConverter(Log.Logger);
                        var goodConvert = converter.ConvertRemarketingFile(config);
                        if (!goodConvert)
                        {
                            HandleError("Conversion failed.");
                        }
                        else
                        {
                            Log.Logger.Information("Conversion successful.");
                        }

                    }
                    catch (Exception e)
                    {
                        HandleError("Runtime Exception", e);
                    }
                })
                .WithNotParsed((errors) =>
                {

                    HandleError("Invalid Configuration");
                    ConverterConfiguration.ShowUsage();
                });

            Console.WriteLine("Press Enter to exit");
            Console.ReadLine();
        }

        static void HandleError(string message, Exception e = null)
        {
            Log.Logger.Error(e, message);

            Environment.ExitCode = -1;
        }

        private static void ConfigureLogger()
        {
            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Information()
                .MinimumLevel.Override("Microsoft", LogEventLevel.Warning)
                .Enrich.FromLogContext()
                .WriteTo.Console()
                .CreateLogger();
        }

        private void ProduceSampleClientConfigurationFile()
        {
            var sampleConfig = new ClientConfiguration
            {
                NumberOfHeaderRows = 3,
                RsaClientId = "1234567890",
                SourceColumnMap = new SourceColumnMap
                {
                    AccountNumber = 1,
                    LoanNumber = 2,
                    LastName = 3,
                    FirstName = 4,
                    LoanBalance = 5,
                    Year = 6,
                    Make = 7,
                    Model = 8,
                    Vin = 9,
                    Mileage = 10,
                    RepoAgentName = 11,
                    RepoAgentsLookup = 12,
                    LocationOfUnit = 13,
                    DateOfRepo = 14,
                    DateOfClear = 15
                }
            };
            var json = JsonConvert.SerializeObject(sampleConfig);
            File.WriteAllText(@"C:\Temp\RMK\clients\sampleClient.json",json);
        }
    }
}
