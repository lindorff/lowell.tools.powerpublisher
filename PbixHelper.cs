using System;
using System.Text;
using System.IO;
using System.IO.Compression;
using CompressionLevel = System.IO.Compression.CompressionLevel;

namespace PowerPublisher
{
    class PbixHelper
    {
        public static void TranslatePbix(string filename, Tuple<string, string>[] translations) //c:/temp/pbix/testrep.pbix
        {
            using (var zipFile = ZipFile.Open(filename, ZipArchiveMode.Update))
            {
                var entry = zipFile.GetEntry("Report/Layout");

                string jsonContent = "";

                using (Stream es = entry.Open())
                {
                    StreamReader sr = new StreamReader(es, System.Text.Encoding.Unicode);
                    jsonContent = sr.ReadToEnd();
                    sr.Close();
                }

                jsonContent = replacePlaceholders(jsonContent, translations);

                entry.Delete();

                ZipArchiveEntry newEntry = zipFile.CreateEntry("Report/Layout", CompressionLevel.Optimal);
                using (Stream es = newEntry.Open())
                {
                    // it's important to have the bigEndian and ByteOrderMark as false
                    using (StreamWriter writer = new StreamWriter(es, new UnicodeEncoding(false, false)))
                    {
                        writer.Write(jsonContent);
                        newEntry.LastWriteTime = DateTimeOffset.UtcNow.LocalDateTime;
                    }
                }

                var sbEntry = zipFile.GetEntry("SecurityBindings");
                sbEntry.Delete();
            }
        }

        private static string replacePlaceholders(string str, Tuple<string, string>[] translations)
        {
            foreach (var t in translations)
            {
                str = str.Replace(t.Item1, t.Item2);
            }
            return str;
        }
    }
}
