/* Title:           Update Design Project
 * Date:            4-1-19
 * Author:          Terry Holmes
 * 
 * Description:     This will update the design project */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using NewEventLogDLL;
using DesignProjectsDLL;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;

namespace WpfApp1
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //setting up the classes
        WPFMessagesClass TheMessagesClass = new WPFMessagesClass();
        EventLogClass TheEventLogClass = new EventLogClass();
        DesignProjectsClass TheDesignProjectsClass = new DesignProjectsClass();

        //setting up the data
        ImportedProjectsDataSet TheImportedProjectsDataSet = new ImportedProjectsDataSet();

        public MainWindow()
        {
            InitializeComponent();
        }

        private void BtnImportExcel_Click(object sender, RoutedEventArgs e)
        {
            Excel.Application xlDropOrder;
            Excel.Workbook xlDropBook;
            Excel.Worksheet xlDropSheet;
            Excel.Range range;

            int intColumnRange = 0;
            int intCounter;
            int intNumberOfRecords;
            string strTransactionID;
            int intTransactionID;
            string strProjectID;
            string strProjectName;
            string strAssignedOffice;
            string strAddress;
            string strCity;
            string strState;
            string strZip;
            

            try
            {
                TheImportedProjectsDataSet.importedprojects.Rows.Clear();

                Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
                dlg.FileName = "Document"; // Default file name
                dlg.DefaultExt = ".xlsx"; // Default file extension
                dlg.Filter = "Excel (.xlsx)|*.xlsx"; // Filter files by extension

                // Show open file dialog box
                Nullable<bool> result = dlg.ShowDialog();

                // Process open file dialog box results
                if (result == true)
                {
                    // Open document
                    string filename = dlg.FileName;
                }

                PleaseWait PleaseWait = new PleaseWait();
                PleaseWait.Show();

                xlDropOrder = new Excel.Application();
                xlDropBook = xlDropOrder.Workbooks.Open(dlg.FileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlDropSheet = (Excel.Worksheet)xlDropOrder.Worksheets.get_Item(1);

                range = xlDropSheet.UsedRange;
                intNumberOfRecords = range.Rows.Count;
                intColumnRange = range.Columns.Count;

                for (intCounter = 1; intCounter <= intNumberOfRecords; intCounter++)
                {
                    strTransactionID = Convert.ToString((range.Cells[intCounter, 1] as Excel.Range).Value2).ToUpper();
                    intTransactionID = Convert.ToInt32(strTransactionID);
                    strProjectID = Convert.ToString((range.Cells[intCounter, 2] as Excel.Range).Value2).ToUpper();
                    strProjectName = Convert.ToString((range.Cells[intCounter, 3] as Excel.Range).Value2).ToUpper();
                    strAssignedOffice = Convert.ToString((range.Cells[intCounter, 4] as Excel.Range).Value2).ToUpper();
                    strAddress = Convert.ToString((range.Cells[intCounter, 5] as Excel.Range).Value2).ToUpper();
                    strCity = Convert.ToString((range.Cells[intCounter, 6] as Excel.Range).Value2).ToUpper();
                    strState = Convert.ToString((range.Cells[intCounter, 7] as Excel.Range).Value2).ToUpper();
                    strZip = Convert.ToString((range.Cells[intCounter, 8] as Excel.Range).Value2).ToUpper();

                    if(strZip != "NULL")
                    {
                        ImportedProjectsDataSet.importedprojectsRow NewProjectRow = TheImportedProjectsDataSet.importedprojects.NewimportedprojectsRow();

                        NewProjectRow.Address = strAddress;
                        NewProjectRow.City = strCity;
                        NewProjectRow.HomeOffice = strAssignedOffice;
                        NewProjectRow.ProjectID = strProjectID;
                        NewProjectRow.ProjectName = strProjectName;
                        NewProjectRow.State = strState;
                        NewProjectRow.TransactionID = intTransactionID;
                        NewProjectRow.Zip = strZip;

                        TheImportedProjectsDataSet.importedprojects.Rows.Add(NewProjectRow);
                    }
                }

                PleaseWait.Close();
                dgrResults.ItemsSource = TheImportedProjectsDataSet.importedprojects;
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "WpfApp1 // Import Excel Button " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        }

        private void BtnProcess_Click(object sender, RoutedEventArgs e)
        {
            int intCounter;
            int intNumberOfRecords;
            int intTransactionID;
            string strState;
            string strZip;
            bool blnFatalError = false;

            try
            {
                intNumberOfRecords = TheImportedProjectsDataSet.importedprojects.Rows.Count - 1;

                for(intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                {
                    intTransactionID = TheImportedProjectsDataSet.importedprojects[intCounter].TransactionID;
                    strState = TheImportedProjectsDataSet.importedprojects[intCounter].State;
                    strZip = TheImportedProjectsDataSet.importedprojects[intCounter].Zip;

                    blnFatalError = TheDesignProjectsClass.UpdateDesignProjectStateZip(intTransactionID, strState, strZip);

                    if (blnFatalError == true)
                        throw new Exception();
                }

                TheMessagesClass.InformationMessage("All Records Have Been Updated");
            }
            catch(Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "WftApp1 // Process Button " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        }
    }
}
