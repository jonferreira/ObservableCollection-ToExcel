using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections.ObjectModel;
//using Ivo_Suite.Classes.Utils;

namespace Classes
{
    public static class Extensions
    {
        /// <summary>
        /// A basic extensions that supports automatic export an ObservableCollection to Excel.
        /// </summary>
        /// <typeparam name="T">The type of elements contained in the collection.</typeparam>
        /// <param name="HiddenProperties">A list of Properties to be hidden</param>
        /// <param name="PropertiesCustomHeader">Defines Custom Header</param>
        /// <param name="ShowExcel">Define if Excel should be displayed</param>
        public static void ToExcel<T>(this ObservableCollection<T> x, List<string> HiddenProperties = null, Dictionary<string, string> PropertiesCustomHeader = null, List<string> SheetCustomHeader = null, bool ShowExcel = true)
        {

            // Load Excel application
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

            // Create empty workbook
            excel.Workbooks.Add();

            // Create Worksheet from active sheet
            Microsoft.Office.Interop.Excel._Worksheet workSheet = excel.ActiveSheet;

            // I created Application and Worksheet objects before try/catch,
            // so that i can close them in finnaly block.
            // It's IMPORTANT to release these COM objects!!

            try
            {

                List<string> propriedades = new List<string>();

                string column = "A";
                int row = 1;

                // ------------------------------------------------
                // Creation of sheet header
                // ------------------------------------------------

                if (SheetCustomHeader != null)
                {
                    foreach (string s in SheetCustomHeader)
                    {
                        workSheet.Cells[row, column] = s;
                        workSheet.Cells[row, 1].EntireRow.Font.Bold = true;
                        workSheet.Range[workSheet.Cells[row, 1], workSheet.Cells[row,x[0].GetType().GetProperties().Count()]].Merge();
                        workSheet.Cells[row, 1].EntireRow.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                        row += 1;
                    }
                    //spacer row
                    row += 1;
                }

                foreach (System.Reflection.PropertyInfo propertyInfo in x[0].GetType().GetProperties())
                {
                    if (propertyInfo.CanRead)
                    {
                        propriedades.Add(propertyInfo.Name);

                        // ------------------------------------------------
                        // Creation of header cells
                        // ------------------------------------------------

                        workSheet.Cells[row, column] = propertyInfo.Name;

                        if (PropertiesCustomHeader != null)
                        {
                            if (PropertiesCustomHeader.ContainsKey(propertyInfo.Name))
                            {
                                workSheet.Cells[row, column] = PropertiesCustomHeader[propertyInfo.Name];
                            }
                        }

                        if (HiddenProperties != null)
                        {
                            if (HiddenProperties.Contains(propertyInfo.Name))
                            {
                                workSheet.Columns[column].Hidden = true;
                            }
                        }

                        column = NextLetter(column);

                    }
                }

                workSheet.Range["A" + row].AutoFormat(Microsoft.Office.Interop.Excel.XlRangeAutoFormat.xlRangeAutoFormatTable1);
                row += 1;

                foreach (T linha in x)
                {
                    column = "A";

                    foreach (string propriedade in propriedades)
                    {
                        workSheet.Cells[row, column] = linha.GetType().GetProperty(propriedade).GetValue(linha, null);
                        column = NextLetter(column);
                    }

                    row += 1;

                }

                // Apply some predefined styles for data to look nicely :)
                //workSheet.Range["A1"].AutoFormat(Microsoft.Office.Interop.Excel.XlRangeAutoFormat.xlRangeAutoFormatClassic1);

                // Define filename
                //string fileName = string.Format(@"{0}\ExcelData.xlsx", Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory));

                // Save this data as a file
                //workSheet.SaveAs(fileName);

                // Display SUCCESS message
                //MessageBox.Show(string.Format("The file '{0}' is saved successfully!", fileName));

            }
            catch (Exception exception)
            {
                //MessageBox.Show("Exception",
                //    "There was a PROBLEM saving Excel file!\n" + exception.Message,
                //    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (!ShowExcel)
                {
                    // Quit Excel application
                    excel.Quit();

                    // Release COM objects (very important!)
                    if (excel != null)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);

                    if (workSheet != null)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(workSheet);

                    // Empty variables
                    excel = null;
                    workSheet = null;

                    // Force garbage collector cleaning
                    GC.Collect();

                }
                else
                {
                    excel.Visible = true;
                }

            }

        }

        public static string NextLetter(string letter, bool MultiChars = true)
        {

            string NextLetter;

            if (MultiChars)
            {
                if (letter == "z")
                {
                    NextLetter = "aa";
                }
                else if (letter == "Z")
                {
                    NextLetter = "AA";
                }
                else
                {
                    NextLetter = ((char)(((int)letter[letter.Length - 1]) + 1)).ToString();
                }

                return NextLetter;
            }
            else
            {
                if (letter == "z")
                {
                    NextLetter = "a";
                }
                else if (letter == "Z")
                {
                    NextLetter = "A";
                }
                else
                {
                    NextLetter = ((char)(((int)letter[letter.Length - 1]) + 1)).ToString();
                }

                return NextLetter;
            }
        }

    }
}
