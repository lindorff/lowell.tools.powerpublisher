using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace PowerPublisher
{
    class TranslationPackage
    {
        private IList<Tuple<string, string, string>> translationData;

        private TranslationPackage(string[] languages, IList<Tuple<string, string, string>> translationData)
        {
            this.Languages = languages;
            this.translationData = translationData;
        }

        public static TranslationPackage Load(string filename)
        {
            IList<Tuple<string, string, string>> translationData = new List<Tuple<string, string, string>>();

            using (var package = new ExcelPackage(new FileInfo(filename)))
            {
                var sheet = package.Workbook.Worksheets[0];

                var languages = GetLanguages(sheet);

                for (int i = 2; i <= sheet.Dimension.End.Row; i++)
                {
                    if (string.IsNullOrEmpty(sheet.Cells[i, 1].Text))
                        break;

                    string key = sheet.Cells[i, 1].Text;

                    for (int j = 2; j < languages.Length+2; j++)
                    {
                        translationData.Add(new Tuple<string, string, string>(key, languages[j - 2], sheet.Cells[i, j].Text));
                    }
                }

                return new TranslationPackage(languages, translationData);
            }
            
        }

        public string GetTranslation(string key, string lang)
        {
            return translationData.Where(p => p.Item1 == key && p.Item2 == lang).Select(p => p.Item3).FirstOrDefault();
        }

        public string[] GetKeys()
        {
            return translationData.Select(p => p.Item1).Distinct().ToArray();
        }

        public string[] Languages
        {
            get;
            private set;
        }

        private static string[] GetLanguages(ExcelWorksheet sheet)
        {
            IList<string> languages = new List<string>();
            for (int i = 2; i <= sheet.Dimension.End.Column; i++)
            {
                if (String.IsNullOrEmpty(sheet.Cells[1, i].Text))
                    break;

                languages.Add(sheet.Cells[1, i].Text);
            }
            return languages.ToArray();
        }
    }
}
