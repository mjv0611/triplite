using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace TrippLite_GUI
{
    //need to also think about if forms move around at all
    //also when an object is empty how to add it.
    //Finally need to figure out how to move information if they just want to transfer it to a different place in the workbook

    //need to check for refresh in the delete function

    //need edit to catch and open a new window if changing places with anothe device used to be
    //If device moves to location that has another device
    //If device is deleted
    //Swap function
    //side note: If multiple select works do I want to be able to turn on multiple device at once
    //need a way to power on/off the internet patches
    //need to fix edit, seems to be grabbing the first entry it finds if the names are the same and you try to change the port information. Change it so mabye it grabs outlets info instead of port 

    public partial class EditForm : Form
    {
        private int arrLoc;
        private string[][] arr;
        [DllImport("user32.dll", SetLastError = true)]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        public EditForm(String[][] xmlArray, string devName, int outlet)
        {
            InitializeComponent();
            arrLoc = -1;
            textBox1.Text = devName;
            for (int i = 0; i < xmlArray.Length - 1; i++)
            {
                arr = xmlArray;
                if (devName == xmlArray[i][0])
                {
                    string ol = xmlArray[i][4];
                    //fills in the names on the treeview1 object.
                    //splits the values of the outlets at the & symbol
                    bool containsOutlet = false;
                    string[] andz = ol.Split('&');
                    if (andz.Length > 1)
                    {
                        //meant for things that hvae more than 2 outlets
                        for (int h = 0; h < andz.Length; h++)
                        {
                            //convert each of the string in the array from string to int to fill in the treeview object correctly
                            int o;
                            bool isOutlet = int.TryParse(andz[h], out o);
                            if (o == outlet)
                                containsOutlet = true;
                        }
                    }
                    else
                    {
                        int o;
                        bool isOutlet = int.TryParse(xmlArray[i][4], out o);
                        if (o == outlet)
                            containsOutlet = true;
                    }
                    if (containsOutlet == true)
                    {
                        textBox2.Text = xmlArray[i][1];
                        textBox3.Text = xmlArray[i][2];
                        textBox4.Text = xmlArray[i][3];
                        string[] ands = xmlArray[i][4].Split('&');
                        if (ands.Length > 1)
                        {
                            //meant for things that hvae more than 2 outlets
                            for (int h = 0; h < ands.Length; h++)
                            {
                                //convert each of the string in the array from string to int to fill in the treeview object correctly
                                int o;
                                bool isOutlet = Int32.TryParse(ands[h], out o);
                                //decides the location of the treeview node to fill in and what to fill in there.
                                if (isOutlet)
                                {
                                    if (h == 0)
                                        comboBox2.SelectedItem = o.ToString();
                                    if (h == 1)
                                        comboBox3.SelectedItem = o.ToString();
                                    if (h == 2)
                                        comboBox4.SelectedItem = o.ToString();
                                    if (h == 3)
                                        comboBox5.SelectedItem = o.ToString();
                                }
                            }
                        }
                        //meant for things with only one outlet
                        else
                        {
                            //o = outlet number
                            int o;
                            bool isOutlet = int.TryParse(xmlArray[i][4], out o);
                            //decides the location of the treeview node to fill in and what to fill in there.
                            if (isOutlet)
                            {
                                comboBox2.SelectedItem = xmlArray[i][4];
                            }
                        }
                        arrLoc = i;
                        refreshCheck();
                    }
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            uint excelId = 0;
            if (arrLoc == -1)
            {
                //do nothing
            }
            else
            {
                Excel.Application ExcelObj = new Excel.Application();
                if (ExcelObj == null)
                {
                    MessageBox.Show("Unable to connect to Microsoft Excel! Terminating program.", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    System.Windows.Forms.Application.Exit();
                }
                Excel.Sheets sheets;
                Excel.Worksheet worksheet;
                //opens the excel book and selects the first sheet
                var books = ExcelObj.Workbooks;
                var theWorkbook = books.Open(@"\\fox\raid\ethernets\Interop\Interop Randomizer\List of Link Partners.xls", 0, false, 5,
                        "", "", true, Excel.XlPlatform.xlWindows, "\t", true, false,
                        0, true);
                //gets all the sheets in the workbook
                sheets = theWorkbook.Worksheets;
                GetWindowThreadProcessId(new IntPtr(ExcelObj.Hwnd), out excelId);
                int sheetNum = -1;
                int rowNum = -1;
                Int32.TryParse(arr[arrLoc][arr[arrLoc].Length - 2], out sheetNum);
                Int32.TryParse(arr[arrLoc][arr[arrLoc].Length - 1], out rowNum);
                worksheet = (Excel.Worksheet)sheets.get_Item(sheetNum);
                var cell = (Excel.Range)worksheet.Cells[ rowNum, "A"];
                cell.Value2 = textBox1.Text;
                cell = (Excel.Range)worksheet.Cells[rowNum, "B"];
                cell.Value2 = textBox2.Text;
                cell = (Excel.Range)worksheet.Cells[rowNum, "C"];
                cell.Value2 = "Rack " + textBox4.Text;
                cell = (Excel.Range)worksheet.Cells[rowNum, "D"];
                cell.Value2 = textBox3.Text;
                cell = (Excel.Range)worksheet.Cells[rowNum, "E"];
                cell.Value2 = textBox4.Text;
                cell = (Excel.Range)worksheet.Cells[rowNum, "F"];
                string outlets = "";
                if (comboBox2.Text != "")
                    outlets = comboBox2.Text;
                if (comboBox3.Text != "")
                {
                    if (outlets == "")
                        outlets = comboBox3.Text;
                    else
                        outlets += " & " + comboBox3.Text;
                }
                if (comboBox4.Text != "")
                {
                    if (outlets == "")
                        outlets = comboBox4.Text;
                    else
                        outlets += " & " + comboBox4.Text;
                }
                if (comboBox5.Text != "")
                {
                    if (outlets == "")
                        outlets = comboBox5.Text;
                    else
                        outlets += " & " + comboBox5.Text;
                }
                cell.Value2 = outlets;
                arr[arrLoc][0] = textBox1.Text;
                arr[arrLoc][1] = textBox2.Text;
                arr[arrLoc][2] = textBox3.Text;
                arr[arrLoc][3] = textBox4.Text;
                arr[arrLoc][4] = outlets;
                theWorkbook.Save();
                //close all excel objects that are currently open
                theWorkbook.Close(false);
                ExcelObj.Quit();
                CloseExcel(sheets);
                CloseExcel(theWorkbook);
                CloseExcel(ExcelObj);
                //the following try catch and process tracking was a solution from Jordy "Kaiwa" Ruiter found on
                //www.codeproject.com/Questions/74980/Close-Excel-Process-with-Interop
                try
                {
                    if (excelId != 0)
                    {
                        Process excel = Process.GetProcessById((int)excelId);
                        excel.CloseMainWindow();
                        excel.Refresh();
                        excel.Kill();
                    }
                }
                catch
                {
                    //process was already killed
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                Close();
            }
        }

        private void refreshCheck( )
        {
            uint excelId = 0;
            if (arrLoc == -1)
            {
                //do nothing
            }
            else
            {
                Excel.Application ExcelObj = new Excel.Application();
                if (ExcelObj == null)
                {
                    MessageBox.Show("Unable to connect to Microsoft Excel! Terminating program.", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    System.Windows.Forms.Application.Exit();
                }
                Excel.Sheets sheets;
                Excel.Worksheet worksheet;
                //opens the excel book and selects the first sheet
                var books = ExcelObj.Workbooks;
                var theWorkbook = books.Open(
                    @"\\fox\raid\ethernets\Interop\Interop Randomizer\List of Link Partners.xls", 0, true, 5,
                    "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false,
                    0, true);
                //gets all the sheets in the workbook
                sheets = theWorkbook.Worksheets;
                GetWindowThreadProcessId(new IntPtr(ExcelObj.Hwnd), out excelId);
                int sheetNum = -1;
                int rowNum = -1;
                Int32.TryParse(arr[arrLoc][arr[arrLoc].Length - 2], out sheetNum);
                Int32.TryParse(arr[arrLoc][arr[arrLoc].Length - 1], out rowNum);
                worksheet = (Excel.Worksheet)sheets.get_Item(sheetNum);
                var cell = (Excel.Range)worksheet.Cells[rowNum, "A"];
                if (cell.Value2 != arr[arrLoc][0])
                {
                    MessageBox.Show("The device information has been changed, please refresh the GUI before proceeding.", "Update Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Load += (s, e) => Close();
                }
                cell = (Excel.Range)worksheet.Cells[rowNum, "B"];
                if (cell.Value2 != null && cell.Value2.ToString() != arr[arrLoc][1])
                {
                    MessageBox.Show("The device information has been changed, please refresh the GUI before proceeding.", "Update Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Load += (s, e) => Close();
                    
                }
                cell = (Excel.Range)worksheet.Cells[rowNum, "D"];
                if (cell.Value2 != arr[arrLoc][2])
                {
                    MessageBox.Show("The device information has been changed, please refresh the GUI before proceeding.", "Update Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Load += (s, e) => Close();
                }
                cell = (Excel.Range)worksheet.Cells[rowNum, "E"];
                if(cell.Value2 != null && cell.Value2.ToString() != arr[arrLoc][3] )
                {
                    MessageBox.Show("The device information has been changed, please refresh the GUI before proceeding.", "Update Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Load += (s, e) => Close();
                }

                //close all excel objects that are currently open
                theWorkbook.Close(false);
                ExcelObj.Quit();
                CloseExcel(sheets);
                CloseExcel(theWorkbook);
                CloseExcel(ExcelObj);
                //the following try catch and process tracking was a solution from Jordy "Kaiwa" Ruiter found on
                //www.codeproject.com/Questions/74980/Close-Excel-Process-with-Interop
                try
                {
                    if (excelId != 0)
                    {
                        Process excel = Process.GetProcessById((int)excelId);
                        excel.CloseMainWindow();
                        excel.Refresh();
                        excel.Kill();
                    }
                }
                catch
                {
                    //process was already killed
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private void CloseExcel(object excel)
        {
            try
            {
                while( System.Runtime.InteropServices.Marshal.ReleaseComObject(excel) > 0 );
            }
            catch (Exception ex)
            {

            }
            finally
            {
                excel = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
        public string[][] updateArray
        {
            get
            {
                return this.arr;
            }
        }
    }
}
