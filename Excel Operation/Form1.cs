using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace Excel_Operation
{
    public partial class Form1 : Form
    {
        DataTable dt;
        public Form1()
        {
            InitializeComponent();
        }
        
        private void loadFile_Click(object sender, EventArgs e)
        {
            int size = -1;
            DialogResult result = openFileDialog1.ShowDialog(); // Show the dialog.

       // C: \Users\sss\Desktop\Story\Test302.xlsx
            if (result == DialogResult.OK) // Test result.
            {
                string file = openFileDialog1.FileName;
                try
                {
                    Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                    Microsoft.Office.Interop.Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(file);
                    Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Sheets[1];
                   // Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet1 = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Sheets[2];

                    Microsoft.Office.Interop.Excel.Range xlRange = xlWorkSheet.UsedRange;
                    int totalRows = xlRange.Rows.Count;
                    int totalColumns = xlRange.Columns.Count;
                    //Microsoft.Office.Interop.Excel.Range xlRange1 = xlWorkSheet1.UsedRange;
                    //int totalRows1 = xlRange1.Rows.Count;
                    //int totalColumns1 = xlRange1.Columns.Count;

                    string query, table;
                    int ordinal;
                    //dgvCurrent.DataSource = dt;
                    for (int rowCount = 2; rowCount <= 265; rowCount++)
                    {
                        textBox1.Text = "";
                        string update = "UPDATE Details SET Ordinal =  {0} WHERE TableName = '{1}' AND GroupId = 3";
                        table = Convert.ToString((xlRange.Cells[rowCount, 1] as Microsoft.Office.Interop.Excel.Range).Text);
                        ordinal = Convert.ToInt16((xlRange.Cells[rowCount, 2] as Microsoft.Office.Interop.Excel.Range).Text);
                        query = string.Format(update, ordinal, table.Trim());

                        xlWorkSheet.Cells[rowCount, 3] = query;

                        textBox1.Text = rowCount.ToString() + " Process";


                        // table1 = Convert.ToString((xlRange1.Cells[rowCount, 1] as Microsoft.Office.Interop.Excel.Range).Text);

                        //for (int row=2; row <= 266; row ++)
                        //{

                        //    table = Convert.ToString((xlRange.Cells[row, 2] as Microsoft.Office.Interop.Excel.Range).Text);
                        //    if (table.ToString().Trim() == table1.ToString().Trim())
                        //    {
                        //        textBox2.Text = "";
                        //        query = Convert.ToString((xlRange.Cells[row, 1] as Microsoft.Office.Interop.Excel.Range).Text);
                        //        table = Convert.ToString((xlRange.Cells[row, 2] as Microsoft.Office.Interop.Excel.Range).Text);
                        //        count = Convert.ToString((xlRange.Cells[row, 3] as Microsoft.Office.Interop.Excel.Range).Text);
                        //        xlWorkSheet1.Cells[rowCount, 2] = table;
                        //        xlWorkSheet1.Cells[rowCount, 5] = query;
                        //        xlWorkSheet1.Cells[rowCount, 3] = count;
                        //        textBox2.Text = row.ToString()+" Added";
                        //    }
                        //    textBox1.Text = rowCount.ToString() + " Process";
                        //}


                        //int startPos = query.LastIndexOf("database") + "[NS_LearnShare].".Length;
                        //int length = query.IndexOf("WHERE") - startPos;
                        //string table = query.Substring(startPos, length);

                        //xlWorkSheet.Cells[rowCount, 3] = query.Length;

                        //xlWorkBook.Close();
                        //xlApp.Quit();
                        //Marshal.ReleaseComObject(xlWorkSheet);
                        //Marshal.ReleaseComObject(xlWorkBook);
                        //Marshal.ReleaseComObject(xlApp);
                    }
                    xlWorkBook.SaveAs(file, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, null, null, null, null, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, null, null, null, null, null);

                    xlWorkBook.Close();
                    xlApp.Quit();

                    Marshal.ReleaseComObject(xlWorkSheet);
                    Marshal.ReleaseComObject(xlWorkBook);
                    Marshal.ReleaseComObject(xlApp);
                }
                catch (IOException)
                {
                }
            }
           


        }
    }
}
