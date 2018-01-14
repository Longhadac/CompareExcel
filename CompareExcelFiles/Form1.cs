using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CompareExcelFiles
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnCompare_Click(object sender, EventArgs e)
        {
            DataTable file1 = ParseExcelFile(txbFile1.Text);
            DataTable file2 = ParseExcelFile(txbFile2.Text);

            //Delete header rows
            int header = int.Parse(txbIgnoreHeaderRow.Text);
            for(int i=0;i<header; i++)
            {
                file1.Rows[i].Delete();
                file2.Rows[i].Delete();
            }
            file1.AcceptChanges();
            file2.AcceptChanges();

            List<int> keys = ParseStringList(txbKeyColumn.Text);
            List<int> compareValues = ParseStringList(txbCompareColumn.Text);
            ExportToExcel(ReverseRowsInDataTable(CompareDataTable(file1, file2,keys,compareValues))
                ,ConfigurationManager.AppSettings.Get("ExportExcelFile"));
        }

        private DataTable ParseExcelFile(string fileName)
        {
            DataTable results = new DataTable();
            string sheetName = ConfigurationManager.AppSettings["SheetName"];

            try
            {
                string connString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 12.0;HDR=no'", fileName);
                string sql = "SELECT * FROM [" + sheetName.ToString() + "]";
                using (OleDbConnection conn = new OleDbConnection(connString))
                {
                    conn.Open();
                    using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                    {
                        using (OleDbDataReader rdr = cmd.ExecuteReader())
                        {
                            results.Load(rdr);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Couldnot parse file");
            }
            return results;
        }

        public DataTable ReverseRowsInDataTable(DataTable inputTable)
        {
            DataTable outputTable = inputTable.Clone();
            for (int i = inputTable.Rows.Count - 1; i >= 0; i--)
            {
                outputTable.ImportRow(inputTable.Rows[i]);
            }
            return outputTable;
        }

        private DataTable CompareDataTable(DataTable file1, DataTable file2, 
            List<int> keys, List<int> compareValues)
        {
            DataTable result = file1.Clone();
            for (int i = file1.Rows.Count - 1; i >= 0; i--)
            {
                DataRow row1 = file1.Rows[i];
                DataRow newRow = row1;
                for (int j = file2.Rows.Count - 1; j >= 0; j--)
                {
                    DataRow row2 = file2.Rows[j];
                    bool check = true;
                    foreach(int key in keys)
                    {
                        if (row1[key].ToString() != row2[key].ToString())
                            check = false;
                    }
                    if(check)
                    {
                        foreach(int k in compareValues)
                        {
                            if (string.IsNullOrEmpty(row1[k].ToString()))
                            {
                                if (!string.IsNullOrEmpty(row2[k].ToString()))
                                {
                                    newRow[k] = "-" + row2[k].ToString();
                                }
                            }
                            else
                            {
                                if (!string.IsNullOrEmpty(row2[k].ToString()))
                                {
                                    newRow[k] = Double.Parse(row1[k].ToString()) - Double.Parse(row2[k].ToString());
                                }
                            }
                        }
                        file2.Rows[j].Delete();
                        file2.AcceptChanges();
                        break;
                    }
                }
                result.Rows.Add(newRow.ItemArray);
                file1.Rows[i].Delete();
                file2.AcceptChanges();
            }

            //Data not found in file1: try to find row in result and insert into list
            //@@TODO: NOT DONE
            if(file2.Rows.Count > 0)
            {
                foreach(DataRow row in file2.Rows)
                {
                    result.Rows.Add(row.ItemArray);
                }
            }
            return result;
        }

        private void ExportToExcel(DataTable DataTable, string ExcelFilePath = null)
        {
            try
            {
                int ColumnsCount;

                if (DataTable == null || (ColumnsCount = DataTable.Columns.Count) == 0)
                    throw new Exception("ExportToExcel: Null or empty input table!\n");

                // load excel, and create a new workbook
                Microsoft.Office.Interop.Excel.Application Excel = new Microsoft.Office.Interop.Excel.Application();
                Excel.Workbooks.Add();

                // single worksheet
                Microsoft.Office.Interop.Excel._Worksheet Worksheet = Excel.ActiveSheet;

                object[] Header = new object[ColumnsCount];

                // column headings               
                for (int i = 0; i < ColumnsCount; i++)
                    Header[i] = DataTable.Columns[i].ColumnName;

                Microsoft.Office.Interop.Excel.Range HeaderRange = Worksheet.get_Range((Microsoft.Office.Interop.Excel.Range)(Worksheet.Cells[1, 1]), (Microsoft.Office.Interop.Excel.Range)(Worksheet.Cells[1, ColumnsCount]));
                HeaderRange.Value = Header;
                HeaderRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                HeaderRange.Font.Bold = true;

                // DataCells
                int RowsCount = DataTable.Rows.Count;
                object[,] Cells = new object[RowsCount, ColumnsCount];

                for (int j = 0; j < RowsCount; j++)
                    for (int i = 0; i < ColumnsCount; i++)
                        Cells[j, i] = DataTable.Rows[j][i];

                Worksheet.get_Range((Microsoft.Office.Interop.Excel.Range)(Worksheet.Cells[2, 1]), (Microsoft.Office.Interop.Excel.Range)(Worksheet.Cells[RowsCount + 1, ColumnsCount])).Value = Cells;

                // check fielpath
                if (ExcelFilePath != null && ExcelFilePath != "")
                {
                    try
                    {
                        Worksheet.SaveAs(ExcelFilePath);
                        Excel.Quit();
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("ExportToExcel: Excel file could not be saved! Check filepath.\n"
                            + ex.Message);
                    }
                }
                else    // no filepath is given
                {
                    Excel.Visible = true;
                }
            }
            catch (Exception ex)
            {
                throw new Exception("ExportToExcel: \n" + ex.Message);
            }
        }

        private void btnFile1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            //openFileDialog1.InitialDirectory = "c:\\";
            openFileDialog1.Filter = "ExcelFiles (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                txbFile1.Text = openFileDialog1.FileName;
            }
        }

        private void btnFile2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            //openFileDialog1.InitialDirectory = "c:\\";
            //openFileDialog1.Filter = "excel files 2003 (*.xls)|*.xls|2007 (*.xlsx)|*.xlsx";
            openFileDialog1.Filter = "ExcelFiles (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                txbFile2.Text = openFileDialog1.FileName;
            }
        }

        private static void WriteLog(string logData, bool logTimeStamp = true)
        {
            try
            {
                using (StreamWriter w = File.AppendText(ConfigurationManager.AppSettings.Get("LogFile")))
                {
                    if (logTimeStamp)
                    {
                        logData = DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToLongTimeString() + ": " + logData;
                    }
                    w.WriteLine(logData);
                }
            }
            catch
            { }
        }

        private List<int> ParseStringList(string list)
        {
            List<int> result = new List<int>();
            if (list.Contains("-"))
            {
                List<string> temp = list.Split(new string[] { "-" },
                   StringSplitOptions.RemoveEmptyEntries).ToList<string>();
                int low = int.Parse(temp[0]);
                int high = int.Parse(temp[1]);
                for(int i=low;i<=high;i++)
                    result.Add(i);
            }
            else
            {
                List<string> temp = list.Split(new string[] { "," }, 
                    StringSplitOptions.RemoveEmptyEntries).ToList<string>();
                result = temp.Select(int.Parse).ToList();
            }
            return result;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            txbIgnoreHeaderRow.Text = ConfigurationManager.AppSettings.Get("IgnoreHeaderRow");
            txbKeyColumn.Text = ConfigurationManager.AppSettings.Get("KeyColumn");
            txbCompareColumn.Text = ConfigurationManager.AppSettings.Get("CompareColumn");
        }
    }
}
