using System;

namespace RmkXlsToXml
{
    public class ClientConfiguration
    {
        public string RsaClientId { get; set; }
        public int NumberOfHeaderRows { get; set; } = 3;
        public SourceColumnMap SourceColumnMap { get; set; }

        public bool Validate()
        {
            if (string.IsNullOrEmpty(RsaClientId))
            {
                Console.WriteLine("Client configuration is missing required RsaClientId value.");
                return false;
            }

            if (NumberOfHeaderRows < 0)
            {
                Console.WriteLine($"Invalid number of header rows configured in Client configuration file. '{NumberOfHeaderRows}'");
                return false;
            }

            if (SourceColumnMap == null)
            {
                Console.WriteLine("Client configuration is missing required Source Column Map.");
                return false;
            }

            if (SourceColumnMap.AccountNumber == null)
            {
                Console.WriteLine("Client Configuration is missing required Account Number column mapping.");
                return false;
            }

            return true;
        }
    }
}