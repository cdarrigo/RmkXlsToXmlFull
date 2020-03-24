using System;
using System.IO;
using CommandLine;
using Newtonsoft.Json;

namespace RmkXlsToXml
{
    public class ConverterConfiguration
    {
        [Option('s', "sourceFile", Required = true, HelpText = "XLS file to convert.")]
        public string SourceFile { get; set; }

        [Option('o', "outputPath", Required = false, HelpText = "Output folder. (default is current folder)", Default = ".")]
        public string OutputPath { get; set; } = ".";

        [Option('c', "ClientConfig", Required = true, HelpText = "Path to client configuration file.")]
        public string ClientConfigFile { get; set; } = ".";

        public ClientConfiguration ClientConfiguration { get; set; }

        public bool IsEmpty()
        {
            return string.IsNullOrEmpty(this.SourceFile) ||
                   string.IsNullOrEmpty(this.OutputPath) ||
                   this.ClientConfiguration == null;
        }

        public bool Validate()
        {
            if (string.IsNullOrEmpty(this.SourceFile))
            {
                Console.WriteLine("Missing required source file argument.");
                return false;
            }

            if (string.IsNullOrEmpty(this.OutputPath))
            {
                Console.WriteLine("Missing required output path argument.");
                return false;
            }

            if (string.IsNullOrEmpty(this.ClientConfigFile))
            {
                Console.WriteLine("Missing required ClientConfig path argument.");
                return false;
            }

            if (!File.Exists(this.SourceFile))
            {
                Console.WriteLine($"Source File '{this.SourceFile}' not found.'");
                return false;
            }

            string absolute = Path.GetFullPath(OutputPath);
            if (string.IsNullOrEmpty(absolute))
            {
                Console.WriteLine($"Invalid output path '{absolute}'.'");
                return false;
            }

            if (!Directory.Exists(absolute))
            {
                try
                {
                    var _ = Directory.CreateDirectory(absolute);
                }
                catch (Exception)
                {
                    Console.WriteLine($"Invalid output path '{absolute}'.'");
                    return false;
                }
            }

            if (!File.Exists(this.ClientConfigFile))
            {
                Console.WriteLine($"Client configuration file '{this.ClientConfigFile}' not found.");
                return false;
            }

            var clientConfigJson = File.ReadAllText(this.ClientConfigFile);
            this.ClientConfiguration = JsonConvert.DeserializeObject<ClientConfiguration>(clientConfigJson);
            return this.ClientConfiguration.Validate();
        }

        public static void ShowUsage()
        {
            string exeName = AppDomain.CurrentDomain.FriendlyName;
            Console.WriteLine($"{exeName} -s <SourceFile> -o <OutputPath> -c <ClientConfigurationFile>");
        }
    }
}