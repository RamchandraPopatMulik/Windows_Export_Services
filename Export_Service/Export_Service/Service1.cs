using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

namespace Export_Service
{
    public partial class Service1 : ServiceBase
    {
        private static Timer aTimer;
        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            aTimer = new Timer(100000);
            aTimer.Elapsed += new ElapsedEventHandler(OnTimedEvent);
            aTimer.Enabled = true;
        }
        private static void OnTimedEvent(object source, ElapsedEventArgs e)
        {
            ExecuteService();
        }
        protected override void OnStop()
        {
            aTimer.Stop();
        }
        private static void ExecuteService()
        {
            DateTime dateTime = DateTime.Now;
            string date = dateTime.ToString("dd_MM_yyyy_hh_mm");
            var file = new FileInfo(@"E:\Sample.xlsx");
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            using (ExcelPackage excel = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = excel.Workbook.Worksheets["Sheet1"];
                SqlConnection connection = new SqlConnection(@"Server=DESKTOP-ICFRQNG;Database=Ram;User Id=DESKTOP-ICFRQNG/Ramchandra;Password=;TrustServerCertificate=True;Integrated Security=SSPI;");
                connection.Open();
                SqlCommand command = new SqlCommand("Select * from Student", connection);
                SqlDataAdapter adapter = new SqlDataAdapter(command);
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);
                worksheet.Cells.LoadFromDataTable(dataTable, true);
                FileInfo excelFile = new FileInfo(@"E:\Basic Core Program\Export_Service\Export" + date + ".xlsx");
                excel.SaveAs(excelFile);
                connection.Close();
            }
        }
    }
}
