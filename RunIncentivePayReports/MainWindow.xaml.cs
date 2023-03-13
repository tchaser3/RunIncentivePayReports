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
using DateSearchDLL;
using DataValidationDLL;
using NewEmployeeDLL;
using NewEventLogDLL;
using IncentivePayDLL;
using EmployeeDateEntryDLL;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System.Windows.Threading;
using System.Timers;

namespace RunIncentivePayReports
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //setting up the classes
        WPFMessagesClass TheMessagesClass = new WPFMessagesClass();
        DateSearchClass TheDateSearchClass = new DateSearchClass();
        DataValidationClass TheDataValidationClass = new DataValidationClass();
        EmployeeClass TheEmployeeClass = new EmployeeClass();
        EventLogClass TheEventLogClass = new EventLogClass();
        IncentivePayClass TheIncentivePayClass = new IncentivePayClass();
        EmployeeDateEntryClass TheEmployeeDateEntryClass = new EmployeeDateEntryClass();
        SendEmailClass TheSendEmailClass = new SendEmailClass();

        //Setting up the DataSets
        FindSortedIncentivePayStatusDataSet TheFindSortedIncentivePayStatusDataSet = new FindSortedIncentivePayStatusDataSet();
        FindIncentivePayByStatusDataSet TheFindIncentivePayByStatusDataSet = new FindIncentivePayByStatusDataSet();
        FindEmployeeByEmployeeIDDataSet TheFindEmployeeByEmployeeIDDataSet = new FindEmployeeByEmployeeIDDataSet();
        FindEmployeeByLastNameDataSet TheFindEmployeeByLastNameDataSet = new FindEmployeeByLastNameDataSet();
        IncentivePayReportsDataSet TheIncentivePayReportsDataSet = new IncentivePayReportsDataSet();

        //setting up global variables
        string gstrUserName;
        string gstrComputerName;
        int gintEmployeeID;
        MessageBoxResult result;
        public static bool gblnAutoRun;

        private static System.Timers.Timer aTimer;

        public MainWindow()
        {
            InitializeComponent();
        }
        private void SetTimer()
        {
            // Create a timer with a two second interval.
            aTimer = new System.Timers.Timer(10000);
            // Hook up the Elapsed event for the timer. 
            aTimer.Elapsed += new ElapsedEventHandler(OnTimedEvent);
            aTimer.AutoReset = true;
            aTimer.Enabled = true;
        }
        private void OnTimedEvent(Object source, ElapsedEventArgs e)
        {
            
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            
            int intCounter;
            int intNumberOfRecords;

            try
            {
                gblnAutoRun = false;

                AutoRunTimer AutoRunTimer = new AutoRunTimer();
                AutoRunTimer.ShowDialog();
                                        
                if(gblnAutoRun == true)
                {                    
                    AutoRunReports();
                }

                gstrComputerName = System.Environment.MachineName;
                gstrUserName = System.Environment.UserName;

                CheckEmployee(gstrComputerName, gstrUserName);

                TheFindSortedIncentivePayStatusDataSet = TheIncentivePayClass.FindSortedIncentivePayStatus();

                intNumberOfRecords = TheFindSortedIncentivePayStatusDataSet.FindSortedIncentivePayStatus.Rows.Count;
                cboSelectStatus.Items.Clear();
                cboSelectStatus.Items.Add("Select Status");

                if(intNumberOfRecords > 0)
                {
                    for(intCounter = 0; intCounter < intNumberOfRecords; intCounter++)
                    {
                        cboSelectStatus.Items.Add(TheFindSortedIncentivePayStatusDataSet.FindSortedIncentivePayStatus[intCounter].TransactionStatus);
                    }

                    cboSelectStatus.SelectedIndex = 0;
                }

                TheIncentivePayReportsDataSet.incentivepayreports.Rows.Clear();

                dgrIncentivePay.ItemsSource = TheIncentivePayReportsDataSet.incentivepayreports;
            }
            catch (Exception Ex)
            {
                TheSendEmailClass.SendEventLog(Ex.ToString());

                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Run Incentive Pay Reports // Main Window // Window Loaded Method " + Ex.ToString());

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        }
        private void CheckEmployee(string strComputerName, string strUserName)
        {
            string strFirstName;
            string strLastName;
            int intCounter;
            int intNumberOfRecords;
            string strTempFirstName;

            try
            {

                strLastName = strUserName.Substring(1).ToUpper();
                strFirstName = strUserName.Substring(0, 1).ToUpper();

                TheFindEmployeeByLastNameDataSet = TheEmployeeClass.FindEmployeesByLastNameKeyWord(strLastName);

                intNumberOfRecords = TheFindEmployeeByLastNameDataSet.FindEmployeeByLastName.Rows.Count;

                if (intNumberOfRecords == 1)
                {
                    gintEmployeeID = TheFindEmployeeByLastNameDataSet.FindEmployeeByLastName[0].EmployeeID;
                }
                else if (intNumberOfRecords > 1)
                {
                    for (intCounter = 0; intCounter < intNumberOfRecords; intCounter++)
                    {
                        strTempFirstName = TheFindEmployeeByLastNameDataSet.FindEmployeeByLastName[intCounter].FirstName.Substring(0, 1).ToUpper();

                        if (strTempFirstName == strFirstName)
                        {
                            gintEmployeeID = TheFindEmployeeByLastNameDataSet.FindEmployeeByLastName[intCounter].EmployeeID;
                        }
                    }
                }

                TheEmployeeDateEntryClass.InsertIntoEmployeeDateEntry(gintEmployeeID, strUserName + " " + strComputerName + " Run Incentive Pay Reports");


            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Run Incentive Pay Reports // Main Window // Check Employee " + Ex.Message);

                TheSendEmailClass.SendEventLog("Run Incentive Pay Reports // Main Window // Check Employee " + Ex.ToString());

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            TheMessagesClass.CloseTheProgram();
        }

        private void cboSelectStatus_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //this will fill up the grid
            int intCounter;
            int intNumberOfRecords;
            int intSelectedIndex;
            int intManagerID;
            int intEmployeeID;
            string strManagerName;
            string strStatus;

            try
            {
                intSelectedIndex = cboSelectStatus.SelectedIndex - 1;

                if(intSelectedIndex > -1)
                {
                    TheIncentivePayReportsDataSet.incentivepayreports.Rows.Clear();
                    strStatus = TheFindSortedIncentivePayStatusDataSet.FindSortedIncentivePayStatus[intSelectedIndex].TransactionStatus;

                    TheFindIncentivePayByStatusDataSet = TheIncentivePayClass.FindIncentivePayByStatus(strStatus);

                    intNumberOfRecords = TheFindIncentivePayByStatusDataSet.FindIncentivePayByStatus.Rows.Count;

                    if(intNumberOfRecords > 0)
                    {
                        for(intCounter = 0; intCounter < intNumberOfRecords; intCounter++)
                        {
                            intEmployeeID = TheFindIncentivePayByStatusDataSet.FindIncentivePayByStatus[intCounter].EmployeeID;

                            TheFindEmployeeByEmployeeIDDataSet = TheEmployeeClass.FindEmployeeByEmployeeID(intEmployeeID);

                            intManagerID = TheFindEmployeeByEmployeeIDDataSet.FindEmployeeByEmployeeID[0].ManagerID;

                            TheFindEmployeeByEmployeeIDDataSet = TheEmployeeClass.FindEmployeeByEmployeeID(intManagerID);

                            strManagerName = TheFindEmployeeByEmployeeIDDataSet.FindEmployeeByEmployeeID[0].FirstName + " ";
                            strManagerName += TheFindEmployeeByEmployeeIDDataSet.FindEmployeeByEmployeeID[0].LastName;

                            IncentivePayReportsDataSet.incentivepayreportsRow NewPayRow = TheIncentivePayReportsDataSet.incentivepayreports.NewincentivepayreportsRow();

                            NewPayRow.AssignedProjectID = TheFindIncentivePayByStatusDataSet.FindIncentivePayByStatus[intCounter].AssignedProjectID;
                            NewPayRow.ProductionDate = TheFindIncentivePayByStatusDataSet.FindIncentivePayByStatus[intCounter].ProductionDate;
                            NewPayRow.CustomerAssignedID = TheFindIncentivePayByStatusDataSet.FindIncentivePayByStatus[intCounter].CustomerAssignedID;
                            NewPayRow.EmployeeName = TheFindIncentivePayByStatusDataSet.FindIncentivePayByStatus[intCounter].Employee;
                            NewPayRow.ManagerName = strManagerName;
                            NewPayRow.PositionTitle = TheFindIncentivePayByStatusDataSet.FindIncentivePayByStatus[intCounter].PositionTitle;
                            NewPayRow.ProjectName = TheFindIncentivePayByStatusDataSet.FindIncentivePayByStatus[intCounter].ProjectName;
                            NewPayRow.RatePerUnit = TheFindIncentivePayByStatusDataSet.FindIncentivePayByStatus[intCounter].RatePerUnit;
                            NewPayRow.TotalIncentivePay = TheFindIncentivePayByStatusDataSet.FindIncentivePayByStatus[intCounter].TotalIncentivePay;
                            NewPayRow.TotalUnits = TheFindIncentivePayByStatusDataSet.FindIncentivePayByStatus[intCounter].TotalUnits;

                            TheIncentivePayReportsDataSet.incentivepayreports.Rows.Add(NewPayRow);
                        }
                    }

                    dgrIncentivePay.ItemsSource = TheIncentivePayReportsDataSet.incentivepayreports;
                }
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Run Incentive Pay Reports // Main Window // cboSelectStatus Selection Changed " + Ex.Message);

                TheSendEmailClass.SendEventLog("Run Incentive Pay Reports // Main Window // cboSelectStatus Selection Changed " + Ex.ToString());

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        }

        private void btnExportToExcel_Click(object sender, RoutedEventArgs e)
        {
            int intRowCounter;
            int intRowNumberOfRecords;
            int intColumnCounter;
            int intColumnNumberOfRecords;

            // Creating a Excel object. 
            Microsoft.Office.Interop.Excel._Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Workbook workbook = excel.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;

            try
            {

                worksheet = workbook.ActiveSheet;

                worksheet.Name = "OpenOrders";

                int cellRowIndex = 1;
                int cellColumnIndex = 1;
                intRowNumberOfRecords = TheIncentivePayReportsDataSet.incentivepayreports.Rows.Count;
                intColumnNumberOfRecords = TheIncentivePayReportsDataSet.incentivepayreports.Columns.Count;

                for (intColumnCounter = 0; intColumnCounter < intColumnNumberOfRecords; intColumnCounter++)
                {
                    worksheet.Cells[cellRowIndex, cellColumnIndex] = TheIncentivePayReportsDataSet.incentivepayreports.Columns[intColumnCounter].ColumnName;

                    cellColumnIndex++;
                }

                cellRowIndex++;
                cellColumnIndex = 1;

                //Loop through each row and read value from each column. 
                for (intRowCounter = 0; intRowCounter < intRowNumberOfRecords; intRowCounter++)
                {
                    for (intColumnCounter = 0; intColumnCounter < intColumnNumberOfRecords; intColumnCounter++)
                    {
                        worksheet.Cells[cellRowIndex, cellColumnIndex] = TheIncentivePayReportsDataSet.incentivepayreports.Rows[intRowCounter][intColumnCounter].ToString();

                        cellColumnIndex++;
                    }
                    cellColumnIndex = 1;
                    cellRowIndex++;
                }

                //Getting the location and file name of the excel to save from user. 
                SaveFileDialog saveDialog = new SaveFileDialog();
                saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                saveDialog.FilterIndex = 1;

                saveDialog.ShowDialog();

                workbook.SaveAs(saveDialog.FileName);
                MessageBox.Show("Export Successful");

            }
            catch (System.Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Run Incentive Pay Reports // Main Window // Export To Excel Button " + Ex.Message);

                TheSendEmailClass.SendEventLog("Run Incentive Pay Reports // Main Window // Export to Excel Button " + Ex.ToString());

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
            finally
            {
                excel.Quit();
                workbook = null;
                excel = null;
            }
        }
        private void AutoRunReports()
        {
            int intStatusCounter;
            int intStatusNumberOfRecords;
            int intPayCounter;
            int intPayNumberOfRecords;
            int intManagerID;
            int intEmployeeID;
            string strManagerName;
            string strStatus;
            int intCounter;
            int intNumberOfRecords;
            string strEmailAddress = "tholmes@bluejaycommunications.com";
            string strHeader = "Incentive Pay Report for " + Convert.ToString(DateTime.Now);
            string strMessage = "";


            try
            {              

                TheFindSortedIncentivePayStatusDataSet = TheIncentivePayClass.FindSortedIncentivePayStatus();

                intStatusNumberOfRecords = TheFindSortedIncentivePayStatusDataSet.FindSortedIncentivePayStatus.Rows.Count;

                for(intStatusCounter = 0; intStatusCounter < intStatusNumberOfRecords; intStatusCounter++)
                {
                    strStatus = TheFindSortedIncentivePayStatusDataSet.FindSortedIncentivePayStatus[intStatusCounter].TransactionStatus;
                    TheIncentivePayReportsDataSet.incentivepayreports.Rows.Clear();

                    if (strStatus != "PAID")
                    {
                        TheFindIncentivePayByStatusDataSet = TheIncentivePayClass.FindIncentivePayByStatus(strStatus);                        

                        intPayNumberOfRecords = TheFindIncentivePayByStatusDataSet.FindIncentivePayByStatus.Rows.Count;

                        if(intPayNumberOfRecords > 0)
                        {
                            for (intPayCounter = 0; intPayCounter < intPayNumberOfRecords; intPayCounter++)
                            {
                                intEmployeeID = TheFindIncentivePayByStatusDataSet.FindIncentivePayByStatus[intPayCounter].EmployeeID;

                                TheFindEmployeeByEmployeeIDDataSet = TheEmployeeClass.FindEmployeeByEmployeeID(intEmployeeID);

                                intManagerID = TheFindEmployeeByEmployeeIDDataSet.FindEmployeeByEmployeeID[0].ManagerID;

                                TheFindEmployeeByEmployeeIDDataSet = TheEmployeeClass.FindEmployeeByEmployeeID(intManagerID);

                                strManagerName = TheFindEmployeeByEmployeeIDDataSet.FindEmployeeByEmployeeID[0].FirstName + " ";
                                strManagerName += TheFindEmployeeByEmployeeIDDataSet.FindEmployeeByEmployeeID[0].LastName;

                                IncentivePayReportsDataSet.incentivepayreportsRow NewPayRow = TheIncentivePayReportsDataSet.incentivepayreports.NewincentivepayreportsRow();

                                NewPayRow.AssignedProjectID = TheFindIncentivePayByStatusDataSet.FindIncentivePayByStatus[intPayCounter].AssignedProjectID;
                                NewPayRow.ProductionDate = TheFindIncentivePayByStatusDataSet.FindIncentivePayByStatus[intPayCounter].ProductionDate;
                                NewPayRow.CustomerAssignedID = TheFindIncentivePayByStatusDataSet.FindIncentivePayByStatus[intPayCounter].CustomerAssignedID;
                                NewPayRow.EmployeeName = TheFindIncentivePayByStatusDataSet.FindIncentivePayByStatus[intPayCounter].Employee;
                                NewPayRow.ManagerName = strManagerName;
                                NewPayRow.PositionTitle = TheFindIncentivePayByStatusDataSet.FindIncentivePayByStatus[intPayCounter].PositionTitle;
                                NewPayRow.ProjectName = TheFindIncentivePayByStatusDataSet.FindIncentivePayByStatus[intPayCounter].ProjectName;
                                NewPayRow.RatePerUnit = TheFindIncentivePayByStatusDataSet.FindIncentivePayByStatus[intPayCounter].RatePerUnit;
                                NewPayRow.TotalIncentivePay = TheFindIncentivePayByStatusDataSet.FindIncentivePayByStatus[intPayCounter].TotalIncentivePay;
                                NewPayRow.TotalUnits = TheFindIncentivePayByStatusDataSet.FindIncentivePayByStatus[intPayCounter].TotalUnits;

                                TheIncentivePayReportsDataSet.incentivepayreports.Rows.Add(NewPayRow);
                            }
                        }
                    }

                    if (TheIncentivePayReportsDataSet.incentivepayreports.Rows.Count > 0)
                    {
                        intNumberOfRecords = TheIncentivePayReportsDataSet.incentivepayreports.Rows.Count;
                        strMessage = "<h1>Incentive Pay Report for Status " + strStatus + "</h1>";
                        strMessage += "<table>";
                        strMessage += "<tr>";
                        strMessage += "<th>Production Date</th>";
                        strMessage += "<th>Employee</th>";
                        strMessage += "<th>Manager</th>";
                        strMessage += "<th>Position Title</th>";
                        strMessage += "<th>Customer Assigned ID</th>";
                        strMessage += "<th>Assigned Project ID</th>";
                        strMessage += "<th>Project Name</th>";
                        strMessage += "<th>Rate Per Unit</th>";
                        strMessage += "<th>Total Units</th>";
                        strMessage += "<th>Total Incentive Pay</th>";
                        strMessage += "</tr>";

                        for (intCounter = 0; intCounter < intNumberOfRecords; intCounter++)
                        {
                            strMessage += "<tr>";
                            strMessage += "<td>" + Convert.ToString(TheIncentivePayReportsDataSet.incentivepayreports[intCounter].ProductionDate)  + "</td>";
                            strMessage += "<td>" + Convert.ToString(TheIncentivePayReportsDataSet.incentivepayreports[intCounter].EmployeeName) + "</td>";
                            strMessage += "<td>" + Convert.ToString(TheIncentivePayReportsDataSet.incentivepayreports[intCounter].ManagerName) + "</td>";
                            strMessage += "<td>" + Convert.ToString(TheIncentivePayReportsDataSet.incentivepayreports[intCounter].PositionTitle) + "</td>";
                            strMessage += "<td>" + Convert.ToString(TheIncentivePayReportsDataSet.incentivepayreports[intCounter].CustomerAssignedID) + "</td>";
                            strMessage += "<td>" + Convert.ToString(TheIncentivePayReportsDataSet.incentivepayreports[intCounter].AssignedProjectID) + "</td>";
                            strMessage += "<td>" + Convert.ToString(TheIncentivePayReportsDataSet.incentivepayreports[intCounter].ProjectName) + "</td>";
                            strMessage += "<td>" + Convert.ToString(TheIncentivePayReportsDataSet.incentivepayreports[intCounter].RatePerUnit) + "</td>";
                            strMessage += "<td>" + Convert.ToString(TheIncentivePayReportsDataSet.incentivepayreports[intCounter].TotalUnits) + "</td>";
                            strMessage += "<td>" + Convert.ToString(TheIncentivePayReportsDataSet.incentivepayreports[intCounter].TotalIncentivePay) + "</td>";
                            strMessage += "</tr>";
                        }

                        strMessage += "</table>";

                        TheSendEmailClass.SendEmail(strEmailAddress, strHeader, strMessage);
                    }
                }                

                //dgrIncentivePay.ItemsSource = TheIncentivePayReportsDataSet.incentivepayreports;
                Application.Current.Shutdown();
            }
            catch(Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Run Incentive Pay Reports // Main Window // Auto Run Reports " + Ex.Message);

                TheSendEmailClass.SendEventLog("Run Incentive Pay Reports // Main Window // Auto Run Reports " + Ex.ToString());

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
            
        }
    }
}
