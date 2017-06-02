using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace FormatedCSV
{
    public partial class Form1 : Form
    {
        public List<Dictionary<string, string>> result;
        public string path;
        public string filename;
        public Form1()
        {
            InitializeComponent();
            BtnConvertSave.Enabled = false;
            result = new List<Dictionary<string, string>>();
        }
        
        private void BtnOpenFile_Click(object sender, EventArgs e)
        {
            List<Dictionary<string, string>> result_temp = new List<Dictionary<string, string>>();
            try
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    //Processing the path and filename
                    string fullpath = openFileDialog1.FileName;
                    textBox1.Text = fullpath;
                    path = Path.GetDirectoryName(fullpath);
                    filename = Path.GetFileNameWithoutExtension(fullpath);

                    Excel.Application xlApp;
                    Excel.Workbook xlWorkBook;
                    Excel.Worksheet xlWorkSheet;
                    object misValue = System.Reflection.Missing.Value;

                    xlApp = new Excel.Application();
                    xlWorkBook = xlApp.Workbooks.Open(openFileDialog1.FileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                    Excel.Range range = xlWorkSheet.UsedRange;
                    int rw = range.Rows.Count;
                    int cl = range.Columns.Count;
                    
                    for (int i = 2; i <= rw; i++)
                    {
                        string streamId = (string)(range.Cells[i, 1] as Excel.Range).Value2;
                        if (string.IsNullOrEmpty(streamId))
                            break;
                        Dictionary<string, string> temp = new Dictionary<string, string>();
                        temp["Stream_ID"] = streamId.TrimStart('0');

                        string fullname = (string)(range.Cells[i, 2] as Excel.Range).Value2;
                        string firstname = string.IsNullOrEmpty(fullname) ? "" : fullname.Split(',')[0];
                        temp["Family_Name"] = firstname;

                        string payroll_date = (string)(range.Cells[i, 5] as Excel.Range).Value2;
                        var formats = new[] { "dd/MM/yyyy", "yyyy-MM-dd" };
                        DateTime dDate;
                        if (DateTime.TryParseExact(payroll_date, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out dDate))
                        {
                            temp["Payroll_Date"] = payroll_date;
                        }
                        else
                        {
                            MessageBox.Show("Invalid Date Format in Line " + i);
                            xlWorkBook.Close(true, misValue, misValue);
                            xlApp.Quit();
                            releaseObject(xlWorkSheet);
                            releaseObject(xlWorkBook);
                            releaseObject(xlApp);
                            System.Windows.Forms.Application.Exit();
                        }

                        double hours = (double)(range.Cells[i, 7] as Excel.Range).Value2;
                        temp["Hours"] = String.Format("{0:0.##}", hours);

                        result_temp.Add(temp);
                    }


                    xlWorkBook.Close(true, misValue, misValue);
                    xlApp.Quit();

                    releaseObject(xlWorkSheet);
                    releaseObject(xlWorkBook);
                    releaseObject(xlApp);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            ///Sort by name, stream_id, payroll_date
            result_temp = result_temp.OrderBy(x=>x["Family_Name"]).ThenBy(x => x["Stream_ID"]).ThenBy(x => x["Payroll_Date"]).ToList();
            
            //Sum up and Sort and Save into List
            if (result_temp!=null && result_temp.Count>1)
            {
                Dictionary<string, string> temp_item = result_temp[0];
                double sum = Convert.ToDouble(result_temp[0]["Hours"]);
                for (int i = 1; i <= result_temp.Count; i++)
                {
                    if (i!=result_temp.Count && result_temp[i]["Stream_ID"] == temp_item["Stream_ID"] && result_temp[i]["Payroll_Date"] == temp_item["Payroll_Date"])
                        sum += Convert.ToDouble(result_temp[i]["Hours"]);
                    else
                    {
                        Dictionary<string, string> ResultLine = new Dictionary<string, string>();
                        ResultLine["Stream_ID"] = result_temp[i - 1]["Stream_ID"];
                        ResultLine["Family_Name"] = result_temp[i - 1]["Family_Name"];
                        ResultLine["Payroll_Date"] = result_temp[i - 1]["Payroll_Date"];
                        ResultLine["Hours"] = String.Format("{0:0.##}", sum);
                        result.Add(ResultLine);

                        if (i != result_temp.Count)
                        {
                            temp_item = result_temp[i];
                            sum = Convert.ToDouble(result_temp[i]["Hours"]);
                        }
                    }

                }
            }

            BtnConvertSave.Enabled = true;
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void BtnConvertSave_Click(object sender, EventArgs e)
        {
            try
            {
                if (!string.IsNullOrEmpty(textBox1.Text))
                {
                    StringBuilder CsvContent = new StringBuilder();
                    string firstLine = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},{19},{20},{21},{22},{23},{24},{25}", "Job UIN", "Run Ref"
                        , "Run Date", "Country", "Location", "Hiring Manager", "First Name"
                        , "Family Name", "Temp Type", "Skillstream ID", "Billing Cost Centre", "Type", "Item Detail", "Rate Name",
                        "Pay Rate", "Agency Rate", "Pay Unit", "Temp Status", "Month", "W/E Date", "Units", "Number of Days Worked", "Net", "Input GST",
                        "Agency Name", "Non-Panel Supplier Name");
                    CsvContent.AppendLine(firstLine);
                    for (int i = 0; i < result.Count; i++)
                    {
                        string newLine = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},{19},{20},{21},{22},{23},{24},{25}", "", "", "", "", "", "", "",
                            result[i]["Family_Name"], "", result[i]["Stream_ID"], "", "Timesheet", "", "base", "", "", "hour", "", "", result[i]["Payroll_Date"], result[i]["Hours"], "", "", "", "", "");
                        CsvContent.AppendLine(newLine);
                    }
                    File.WriteAllText(path+filename+".csv", CsvContent.ToString());
                    MessageBox.Show("Saved Successfully!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            
        }
    }
}
