using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LocalizationTool
{
    class Program
    {
        private static List<string> m_langIDs = new List<string>();
        private static List<string> m_textIDs = new List<string>();
        private static List<List<string>> m_textData = new List<List<string>>();

        static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: <app.exe> <excel file path> <code file path>.");
                return;
            }

            var excelFile = args[0];
            var codeFile = args[1];

            using (var stream = File.Open(excelFile, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var dataSet = reader.AsDataSet();

                    ReadExcelTable(dataSet.Tables[0]);
                    GenerateCodeFile(codeFile);

                    Console.WriteLine("Done with success.");
                }
            }
        }

        static void ReadExcelTable(DataTable table)
        {
            for (var i = 0; i < table.Rows.Count; i++)
            {
                var row = table.Rows[i];

                // Language IDs.
                if (i == 0)
                {
                    var items = row.ItemArray;
                    for (var j = 1; j < items.Length; j++)
                    {
                        m_langIDs.Add(items[j] as string);
                    }
                }
                else
                {
                    var tempArr = new List<string>();
                    var items = row.ItemArray;
                    for (var j = 0; j < items.Length; j++)
                    {
                        if (j == 0)
                        {
                            // Text ID.
                            m_textIDs.Add(items[j] as string);
                        }
                        else
                        {
                            // Text data.
                            tempArr.Add(items[j] as string);
                        }
                    }

                    m_textData.Add(tempArr);
                }
            }
        }

        static string JoinTexts(List<string> texts, bool hasQuote)
        {
            var joinedText = "";
            var n = texts.Count;
            for (var i = 0; i < n; i++)
            {
                var text = hasQuote ? "\"" + texts[i] + "\"" : texts[i];
                joinedText += text;
                if (i < n - 1)
                {
                    joinedText += ", ";
                }
            }

            return joinedText;
        }

        static string JoinTexts(List<List<string>> texts, int nSpaces)
        {
            var spacesText = "";
            for (var i = 0; i < nSpaces; i++)
            {
                spacesText += " ";
            }

            var joinedText = "";
            var n = texts.Count;
            for (var i = 0; i < n; i++)
            {
                var temp = string.Format("{0}{{ {1} }}", spacesText, JoinTexts(texts[i], true));
                if (i < n - 1)
                {
                    temp += ",\n";
                }

                joinedText += temp;
            }

            return joinedText;
        }

        static void GenerateCodeFile(string filename)
        {
            var langIDsText = JoinTexts(m_langIDs, false);
            var textIDsText = JoinTexts(m_textIDs, false);
            var textDataText = JoinTexts(m_textData, 8);

            var text = string.Format(@"public class LocalizationData
{{
    public enum LangID {{ {0} }}
    public enum TextID {{ {1} }}

    private static string[,] m_data = 
    {{
{2}
    }};

    public static string GetText(LangID langID, TextID textID)
    {{
        return m_data[(int)textID, (int)langID];
    }}
}}", langIDsText, textIDsText, textDataText);

            File.WriteAllText(filename, text);
        }
    }
}
