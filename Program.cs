using Microsoft.Extensions.Configuration;
using System;
using System.IO;
using System.Linq;

namespace PowerPublisher
{
    class Program
    {
        static void Main(string[] args)
        {
            ValidateParameters(args);

            var pbixFile = args[1];
            var translationFile = args[2];

            TranslationPackage pckg = null;
            try
            {
                pckg = TranslationPackage.Load(translationFile);
            } 
            catch
            {
                Console.Write("Unable to read translations file!");
                return;
            }

            try
            {
                foreach (var lang in pckg.Languages)
                {
                    string outputFile = GetVersionedReportFilename(args[1], lang);
                    File.Copy(pbixFile, outputFile, true);
                    var translations = pckg.GetKeys().Select(p => new Tuple<string, string>(p, pckg.GetTranslation(p, lang)));
                    PbixHelper.TranslatePbix(outputFile, translations.ToArray());
                }
            } 
            catch
            {
                Console.Write("Unable to create translated pbix files!");
                return;
            }

            if (args[0] == "translate-publish")
            {
                PowerBiService pbis = new PowerBiService(ReadSettings());

                string reportName = Path.GetFileName(pbixFile);

                foreach (var lang in pckg.Languages)
                {
                    pbis.PublishReport(GetVersionedReportFilename(pbixFile, lang), GetVersionedReportFilename(reportName, lang));
                }
            }

        }

        private static void ValidateParameters(string[] args)
        {
            if (args.Length < 3 || (args[0] != "translate" && args[0] != "translate-publish"))
            {
                PrintUsage();
                System.Environment.Exit(0);
            }

            if (!File.Exists(args[1]))
            {
                Console.WriteLine("Input file {0} not found!", args[1]);
                System.Environment.Exit(-11);
            }

            if (!File.Exists(args[2]))
            {
                Console.WriteLine("Input file {0} not found!", args[2]);
                System.Environment.Exit(-12);
            }
        }

        private static void PrintUsage()
        {
            Console.WriteLine("Only create local translated copies:");
            Console.WriteLine("    PowerPublisher translate <inputFile.pbix> <translations.xlsx>");
            Console.WriteLine("Create local translated copies and publish to PowerBI service:");
            Console.WriteLine("    PowerPublisher translate-publish <inputFile.pbix> <translations.xlsx>");
        }

        private static PowerBiSettings ReadSettings()
        {
            IConfiguration config = new ConfigurationBuilder()
                .AddJsonFile("appsettings.json", true, true)
                .Build();

            return config.GetSection("PowerBiApi").Get<PowerBiSettings>();
        }

        private static string GetVersionedReportFilename(string basePbixFile, string lang)
        {
            return basePbixFile.Replace(".pbix", "-" + lang + ".pbix");
        }
    }
}
