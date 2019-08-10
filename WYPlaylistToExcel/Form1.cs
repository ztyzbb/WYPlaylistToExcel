using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Newtonsoft.Json.Linq;
using System;
using System.Windows.Forms;

namespace WYPlaylistToExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void ExportToExcel(object sender, EventArgs e)
        {
            JObject json = JObject.Parse(textBox1.Text);
            JObject playlist = (JObject)json["playlist"];

            // 创建文档
            // By default, AutoSave = true, Editable = true, and Type = xlsx.
            SpreadsheetDocument document = SpreadsheetDocument.Create("out.xlsx", SpreadsheetDocumentType.Workbook);

            // 准备文件结构中的各XML
            // Add a WorkbookPart to the document.
            WorkbookPart workbookPart = document.AddWorkbookPart();
            // Add a SharedStringTablePart to the WorkbookPart.
            SharedStringTablePart sharedStringTablePart = workbookPart.AddNewPart<SharedStringTablePart>();
            // Add a WorksheetPart to the WorkbookPart.
            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

            // 生成WorkbookPart的XML结构
            Workbook workbook = workbookPart.Workbook = new Workbook();
            // Add Sheets to the Workbook.
            Sheets sheets = workbook.AppendChild(new Sheets());
            // Append a new worksheet and associate it with the workbook.
            sheets.Append(new Sheet()
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = (string)playlist["name"]
            });

            // 生成SharedStringTablePart的XML结构
            SharedStringTable sharedStringTable = sharedStringTablePart.SharedStringTable = new SharedStringTable();

            // 生成WorksheetPart的XML结构
            Worksheet worksheet = worksheetPart.Worksheet = new Worksheet();
            // Get the sheetData cell table.
            SheetData sheetData = worksheet.AppendChild(new SheetData());

            // 插入数据
            AddDatas(sheetData, sharedStringTable, (JArray)playlist["tracks"]);

            // 关闭文档
            document.Close();
            MessageBox.Show("导出成功！");
        }

        private void AddDatas(SheetData sheetData, SharedStringTable sharedStringTable, JArray tracks)
        {
            uint count = 1;
            foreach (JObject track in tracks)
            {
                // Add a row to the cell table.
                Row row = new Row() { RowIndex = count };
                sheetData.Append(row);
                string countStr = count.ToString();

                row.AppendChild(new Cell()
                {
                    CellReference = "A" + countStr,
                    CellValue = new CellValue((string)track["name"]),
                    DataType = new EnumValue<CellValues>(CellValues.String)
                });

                string artistStr = string.Empty;
                foreach (JObject artist in track["ar"])
                {
                    artistStr += (string)artist["name"] + ';';
                }

                row.AppendChild(new Cell()
                {
                    CellReference = "B" + countStr,
                    CellValue = new CellValue(InsertSharedStringItem(artistStr.Substring(0, artistStr.Length - 1), sharedStringTable).ToString()),
                    DataType = new EnumValue<CellValues>(CellValues.SharedString)
                });

                row.AppendChild(new Cell()
                {
                    CellReference = "C" + countStr,
                    CellValue = new CellValue(InsertSharedStringItem((string)track["al"]["name"], sharedStringTable).ToString()),
                    DataType = new EnumValue<CellValues>(CellValues.SharedString)
                });

                uint time = (uint)track["dt"];
                row.AppendChild(new Cell()
                {
                    CellReference = "D" + countStr,
                    CellValue = new CellValue($"{time / 60000:00}:{time % 60000 / 1000:00}.{time % 1000:000}"),
                    DataType = new EnumValue<CellValues>(CellValues.String)
                });

                count++;
            }
        }

        private char[] chars = { 'Z', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y' };

        private string NumberToAAA(uint number)
        {
            string result = string.Empty;
            while (number > 0)
            {
                result = chars[number % 26] + result;
                number = (number - 1) / 26;
            }
            return result;
        }

        private int InsertSharedStringItem(string text, SharedStringTable sharedStringTable)
        {
            int i = 0;

            // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
            foreach (SharedStringItem item in sharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    return i;
                }

                i++;
            }

            // The text does not exist in the part. Create the SharedStringItem and return its index.
            sharedStringTable.AppendChild(new SharedStringItem(new Text(text)));

            return i;
        }
    }
}
