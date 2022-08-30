using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Reflection;
using Microsoft.Office.Interop.Excel;

namespace ExcelAutomation
{
    public partial class ExcelAutomation : Form
    {
        Excel.Application objXL;
        Excel._Workbook objBook;

        public ExcelAutomation()
        {
            InitializeComponent();
        }

        private void launch_excel(object sender, EventArgs e)
        {
            
            try
            {
                //Start Excel
                objXL = new Excel.Application();
                objXL.Visible = true;

                //Create new workbook
                objBook = (Excel._Workbook)(objXL.Workbooks.Add(Missing.Value));

                objXL.Visible = true;
                objXL.UserControl = true;
            }
            catch (Exception theException)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);

                MessageBox.Show(errorMessage, "Error");
            }
        }

        private void select_cell(object sender, EventArgs e)
        {
            Excel.Sheets objSheets;
            Excel._Worksheet objSheet;
            Excel.Range range;

            try
            {
                objSheets = objBook.Worksheets;
                objSheet = (Excel._Worksheet)objSheets.get_Item(1);

                objSheet.Cells[1, 1] = "First Name";
            }
            catch(Exception theException)
            {
                String errorMessage;
                errorMessage = "Can't find the Excel workbook.";

                MessageBox.Show(errorMessage, "Missing Workbook?");

                //You can't automate Excel if you can't find the data you created, so 
                //leave the subroutine.
                return;
            }

        }
    }

}
