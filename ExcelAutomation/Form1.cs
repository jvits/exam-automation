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
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Status;

namespace ExcelAutomation
{
    public partial class ExcelAutomation : Form
    {
        Excel.Application objXL;
        Excel._Workbook objBook;

        String errorMessage;
        

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

        private void update_cell(object sender, EventArgs e)
        {
            Excel.Sheets objSheets;
            Excel._Worksheet objSheet;

            try
            {
                //Value from selected cells text box
                string _selectedCell;
                int _selectedCellValue;

                string[] _selectedCells;

                _selectedCell = selectedCell.Text;
                _selectedCells = _selectedCell.Split('-');

                //Cell value validation
                try
                {
                    _selectedCellValue = Int32.Parse(selectedCellValue.Text);
                }
                catch (Exception theException)
                {

                    errorMessage = "Invalid cell value.";
                    MessageBox.Show(errorMessage, "Missing cell?");
                    return;

                }
                
                

            
                objSheets = objBook.Worksheets;
                objSheet = (Excel._Worksheet)objSheets.get_Item(1);
                if(_selectedCell.Length <= 2)
                {
                    objSheet.get_Range(_selectedCells[0], _selectedCells[0]).Value2 = _selectedCellValue;
                }
                else
                {
                    objSheet.get_Range(_selectedCells[0], _selectedCells[1]).Value2 = _selectedCellValue;
                }
            }
            catch(Exception theException)
            {

                errorMessage = "Invalid cell range.";
                MessageBox.Show(errorMessage, "Missing cell?");
                return;

            }

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void submitFormula_Click(object sender, EventArgs e)
        {
            Excel.Sheets objSheets;
            Excel._Worksheet objSheet;
            Excel.Range oRng;

            try
            {
                //Value from selected cells text box formula
                string _selectedCellResult;
                string _selectedCellFormula;

                string[] _selectedCells;

                _selectedCellResult = selectedCellResult.Text;
                _selectedCellFormula = selectedCellFormula.Text;
                _selectedCells = _selectedCellResult.Split('-');

                objSheets = objBook.Worksheets;
                objSheet = (Excel._Worksheet)objSheets.get_Item(1);
                if (_selectedCells.Length <= 2)
                {
                    oRng = objSheet.get_Range(_selectedCells[0], _selectedCells[0]);
                    oRng.Formula = _selectedCellFormula;
                }
                else
                {
                    oRng = objSheet.get_Range(_selectedCells[0], _selectedCells[1]);
                    oRng.Formula = _selectedCellFormula;
                    
                }

                displayFormula.Text = "Formula:" + oRng.Formula.ToString();
                displayFormula.Refresh();

                //Evaluate result value of formula per cell
                //string evalFormula = objXL.Evaluate(_selectedCellFormula).ToString();

            }
            catch(Exception theException)
            {
                errorMessage = theException.ToString();
                MessageBox.Show(errorMessage, "Missing cell?");
                return;
            }

        }

        private void selectedCellResult_TextChanged(object sender, EventArgs e)
        {
            Excel.Sheets objSheets;
            Excel._Worksheet objSheet;
            Excel.Range oRng;

            try
            {
                //Value from selected cells text box formula
                string _selectedCellResult;
                string[] _selectedCells;

                _selectedCellResult = selectedCellResult.Text;
                _selectedCells = _selectedCellResult.Split('-');

                objSheets = objBook.Worksheets;
                objSheet = (Excel._Worksheet)objSheets.get_Item(1);
                if (_selectedCells.Length <= 2)
                {
                    oRng = objSheet.get_Range(_selectedCells[0], _selectedCells[0]);
                }
                else
                {
                    oRng = objSheet.get_Range(_selectedCells[0], _selectedCells[1]);

                }

               

                displayFormula.Text = "Formula: " + oRng.Formula.ToString();
                displayFormula.Refresh();

            }
            catch(Exception theException)
            {
                //errorMessage = theException.ToString();
                //MessageBox.Show(errorMessage, "Missing cell?");
                return;
            }
        }
    }

}
