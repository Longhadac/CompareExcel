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
            for(int i=0;i<13;i++)
            {
                file1.Rows[i].Delete();
                file2.Rows[i].Delete();
            }
            file1.AcceptChanges();
            file2.AcceptChanges();            
            
            ExportToExcel(ReverseRowsInDataTable(CompareDataTable(file1, file2))
                , ConfigurationManager.AppSettings.Get("ExportExcelFile"));
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
            List<string> keys, List<string> compareValues)
        {
            DataTable result = file1.Clone();
            for (int i = file1.Rows.Count - 1; i >= 0; i--)
            {
                DataRow row1 = file1.Rows[i];
                DataRow newRow = row1;
                for (int j = file2.Rows.Count - 1; j >= 0; j--)
                {
                    DataRow row2 = file2.Rows[j];
                    if ((row1[1].ToString() == row2[1].ToString())
                        && (row1[3].ToString() == row2[3].ToString())
                        && (row1[4].ToString() == row2[4].ToString())
                        && (row1[5].ToString() == row2[5].ToString()))
                    {
                        for (int k = 6; k < 35; k++)
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

        private List<string> ParseStringList(string list)
        {
            return list.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries).ToList<string>();
        }   
    }
}
