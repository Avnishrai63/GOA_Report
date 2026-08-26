using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelDataReader;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Timers;

namespace GOA_Report
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            this.MaximizeBox = false;

            timer1.Interval = 15;
            timer1.Tick += timer1_Tick;

            label6.Text = DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt");

            timer1.Start();


            //timer1.Interval = 1000;
            //timer1.Tick += timer1_Tick;
            //timer1.Start();

            //// Pehli baar turant date/time show kare
            //timer1_Tick(null, null);


            //label6.Text = DateTime.Now.ToString("dd/MM/yyyy           hh:mm:ss tt");

            //this.FormBorderStyle = FormBorderStyle.FixedDialog;

        }

        List<ReportData> data = new List<ReportData>();


        private void button1_Click(object sender, EventArgs e)
        {
            data.Clear();
            string sourceFile = @"\\192.168.0.223\Production Share\GDP207_Feedback\Report.xls";

            using (var stream = new FileStream(
     sourceFile,
     FileMode.Open,
     FileAccess.Read,
     FileShare.ReadWrite | FileShare.Delete))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet();

                    DataTable dt = result.Tables[0];   // First Sheet

                    for (int i = 1; i < dt.Rows.Count; i++) // Header skip
                    {
                        DataRow row = dt.Rows[i];

                        DateTime date;
                        DateTime.TryParse(row[1].ToString(), out date);

                        int korr = 0;
                        int.TryParse(row[6].ToString(), out korr);

                        double time = 0;
                        double.TryParse(row[7].ToString(), out time);

                        data.Add(new ReportData
                        {
                            Date = date,
                            User1_M = row[3].ToString().Trim(),          // User1(M)
                            US1_Prs1_M = row[4].ToString().Trim(),       // GOZ
                            Korrigiertefelder = korr,                    // Column G
                            Time = time                                  // Column H
                        });
                    }
                }
            }

            int month = dateTimePicker1.Value.Month;
            int year = dateTimePicker1.Value.Year;

            var report = data
    .Where(x => !string.IsNullOrWhiteSpace(x.User1_M)
             && x.US1_Prs1_M == "GOZ"
             && x.Date.Year == year
             && x.Date.Month == month)
    .GroupBy(x => new
    {
        Month = x.Date.ToString("yyyy-MM"),
        User = x.User1_M
    })
    .Select(g => new
    {
        Month = g.Key.Month,
        User = g.Key.User,
        TotalRecord = g.Count(),
        TotalKorrigiertefelder = g.Sum(x => x.Korrigiertefelder),
        TotalTime = g.Sum(x => x.Time)
    })
    .ToList();

            if (report.Count == 0)
            {
                MessageBox.Show("GOZ data not found.");
                return;
            }

            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Report");

            // Header
            ws.Cell(1, 1).Value = "Month";
            ws.Cell(1, 2).Value = "User";
            ws.Cell(1, 3).Value = "Korrigiertefelder";
            ws.Cell(1, 4).Value = "Time";
            ws.Cell(1, 5).Value = "PerHours";

            int row1 = 2;

            foreach (var item in report)
            {
                ws.Cell(row1, 1).Value = item.Month;
                ws.Cell(row1, 2).Value = item.User;
                ws.Cell(row1, 3).Value = item.TotalKorrigiertefelder;
                ws.Cell(row1, 4).Value = item.TotalTime;
                double perHours = 0;

                if (item.TotalTime != 0)
                {
                    perHours = Convert.ToDouble(item.TotalKorrigiertefelder) / Convert.ToDouble(item.TotalTime);
                }

                ws.Cell(row1, 5).Value = perHours;
                row1++;
            }

            //// Column width auto fit
            //ws.Columns().AdjustToContents();

            //// Save
            //wb.SaveAs(@"D:\Report1.xlsx");
            //MessageBox.Show("GOZ Report created succesfully...");
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            // Folder Name
            string folderPath = Path.Combine(desktopPath, "GOZ_Reports");

            // Folder create if not exists
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }

            // Agar report ek hi month ka hai
            string monthName = report.First().Month;   // e.g. 2026-04

            string filePath = Path.Combine(folderPath, "GOZ_Report_" + month + "_" + year + ".xlsx");

            // Save Excel
            wb.SaveAs(filePath);

            MessageBox.Show("Report created successfully.....");

        }

        private void button2_Click(object sender, EventArgs e)
        {


        }

        private void button3_Click(object sender, EventArgs e)
        {


        }

        private void button4_Click(object sender, EventArgs e)
        {
            data.Clear();

            string sourceFile = @"\\192.168.0.223\Production Share\GDP207_Feedback\Report.xls";

            using (var stream = new FileStream(
     sourceFile,
     FileMode.Open,
     FileAccess.Read,
     FileShare.ReadWrite | FileShare.Delete))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet();

                    DataTable dt = result.Tables[0];




                    //data.Clear();

                    for (int i = 1; i < dt.Rows.Count; i++)
                    {
                        DataRow row = dt.Rows[i];

                        // User2 blank hai to row skip
                        if (string.IsNullOrWhiteSpace(row[12].ToString()))
                            continue;

                        DateTime date;
                        DateTime.TryParse(row[1].ToString(), out date);

                        int korr = 0;
                        int.TryParse(row[15].ToString(), out korr);

                        double time = 0;
                        double.TryParse(row[16].ToString(), out time);

                        data.Add(new ReportData
                        {
                            Date = date,
                            User1_M = row[12].ToString().Trim(),
                            US1_Prs1_M = row[13].ToString().Trim(),
                            Korrigiertefelder = korr,
                            Time = time
                        });
                    }

                    //foreach (var item in data.Take(10))
                    //{
                    //    MessageBox.Show(
                    //        "Date = " + item.Date +
                    //        "\nUser = " + item.User1_M +
                    //        "\nProcess = " + item.US1_Prs1_M +
                    //        "\nKorr = " + item.Korrigiertefelder +
                    //        "\nTime = " + item.Time);
                    //}

                    int month = dateTimePicker2.Value.Month;
                    int year = dateTimePicker2.Value.Year;

                    var report = data
        .Where(x => !string.IsNullOrWhiteSpace(x.User1_M)
                 && x.US1_Prs1_M == "GOA"
                 && x.Date.Year == year
                 && x.Date.Month == month)
        .GroupBy(x => new
        {
            Month = x.Date.ToString("yyyy-MM"),
            User = x.User1_M
        })
        .Select(g => new
        {
            Month = g.Key.Month,
            User = g.Key.User,
            TotalKorrigiertefelder = g.Sum(x => x.Korrigiertefelder),
            TotalTime = g.Sum(x => x.Time)
        })
        .ToList();

                    if (report.Count == 0)
                    {
                        MessageBox.Show("GOA data not found.");
                        return;
                    }

                    var wb = new XLWorkbook();
                    var ws = wb.Worksheets.Add("GOA Report");

                    // Header
                    ws.Cell(1, 1).Value = "Month";
                    ws.Cell(1, 2).Value = "User";
                    ws.Cell(1, 3).Value = "Korrigiertefelder";
                    ws.Cell(1, 4).Value = "Time";
                    ws.Cell(1, 5).Value = "PerHours";

                    int row1 = 2;

                    foreach (var item in report)
                    {
                        ws.Cell(row1, 1).Value = item.Month;
                        ws.Cell(row1, 2).Value = item.User;
                        ws.Cell(row1, 3).Value = item.TotalKorrigiertefelder;
                        ws.Cell(row1, 4).Value = item.TotalTime;
                        double perHours = 0;

                        if (item.TotalTime != 0)
                        {
                            perHours = Convert.ToDouble(item.TotalKorrigiertefelder) / Convert.ToDouble(item.TotalTime);
                        }

                        ws.Cell(row1, 5).Value = perHours;

                        row1++;
                    }

                    ws.Columns().AdjustToContents();

                    string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                    string folderPath = Path.Combine(desktopPath, "GOA_Reports");

                    if (!Directory.Exists(folderPath))
                    {
                        Directory.CreateDirectory(folderPath);
                    }

                    string filePath = Path.Combine(folderPath, "GOA_Report_" + month + "_" + year + ".xlsx");

                    wb.SaveAs(filePath);

                    MessageBox.Show("GOA Report created successfully...");
                }
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            data.Clear();

            string sourceFile = @"\\192.168.0.223\Production Share\GDP207_Feedback\Report.xls";

            using (var stream = new FileStream(
     sourceFile,
     FileMode.Open,
     FileAccess.Read,
     FileShare.ReadWrite | FileShare.Delete))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet();

                    DataTable dt = result.Tables[0];




                    data.Clear();

                    for (int i = 1; i < dt.Rows.Count; i++)
                    {
                        DataRow row = dt.Rows[i];

                        if (string.IsNullOrWhiteSpace(row[21].ToString()))
                            continue;

                        DateTime date;
                        DateTime.TryParse(row[1].ToString(), out date);

                        int korr = 0;
                        int.TryParse(row[24].ToString(), out korr);

                        double time = 0;
                        double.TryParse(row[25].ToString(), out time);

                        data.Add(new ReportData
                        {
                            Date = date,
                            User1_M = row[21].ToString().Trim(),
                            US1_Prs1_M = row[22].ToString().Trim(),
                            Korrigiertefelder = korr,
                            Time = time
                        });
                    }

                    int month = dateTimePicker3.Value.Month;
                    int year = dateTimePicker3.Value.Year;

                    var report = data
                    .Where(x => !string.IsNullOrWhiteSpace(x.User1_M)
                             && x.US1_Prs1_M == "NK1"
                             && x.Date.Year == year
                             && x.Date.Month == month)
                    .GroupBy(x => new
                    {
                        Month = x.Date.ToString("yyyy-MM"),
                        User = x.User1_M
                    }) 
                    .Select(g => new
                    {
                        Month = g.Key.Month,
                        User = g.Key.User,
                        TotalKorrigiertefelder = g.Sum(x => x.Korrigiertefelder),
                        TotalTime = g.Sum(x => x.Time)
                    })
                    .ToList();

                    if (report.Count == 0)
                    {
                        MessageBox.Show("NK1 data not found.");
                        return;
                    }

                    var wb = new XLWorkbook();
                    var ws = wb.Worksheets.Add("NK1 Report");

                    // Header
                    ws.Cell(1, 1).Value = "Month";
                    ws.Cell(1, 2).Value = "User";
                    ws.Cell(1, 3).Value = "Korrigiertefelder";
                    ws.Cell(1, 4).Value = "Time";
                    ws.Cell(1, 5).Value = "PerHours";

                    int row1 = 2;

                    foreach (var item in report)
                    {
                        ws.Cell(row1, 1).Value = item.Month;
                        ws.Cell(row1, 2).Value = item.User;
                        ws.Cell(row1, 3).Value = item.TotalKorrigiertefelder;
                        ws.Cell(row1, 4).Value = item.TotalTime;
                        double perHours = 0;

                        if (item.TotalTime != 0)
                        {
                            perHours = Convert.ToDouble(item.TotalKorrigiertefelder) / Convert.ToDouble(item.TotalTime);
                        }

                        ws.Cell(row1, 5).Value = perHours;

                        row1++;
                    }

                    ws.Columns().AdjustToContents();

                    string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                    string folderPath = Path.Combine(desktopPath, "NK1_Reports");

                    if (!Directory.Exists(folderPath))
                    {
                        Directory.CreateDirectory(folderPath);
                    }

                    string filePath = Path.Combine(folderPath, "NK1_Report_" + month + "_" + year + ".xlsx");

                    wb.SaveAs(filePath);

                    MessageBox.Show("NK1 Report created successfully...");
                }
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            data.Clear();

            string sourceFile = @"\\192.168.0.223\Production Share\GDP207_Feedback\Report.xls";

            using (var stream = new FileStream(
     sourceFile,
     FileMode.Open,
     FileAccess.Read,
     FileShare.ReadWrite | FileShare.Delete))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet();

                    DataTable dt = result.Tables[0];



                    data.Clear();

                    for (int i = 1; i < dt.Rows.Count; i++)
                    {
                        DataRow row = dt.Rows[i];

                        if (string.IsNullOrWhiteSpace(row[30].ToString()))
                            continue;

                        DateTime date;
                        DateTime.TryParse(row[1].ToString(), out date);

                        int korr = 0;
                        int.TryParse(row[33].ToString(), out korr);

                        double time = 0;
                        double.TryParse(row[34].ToString(), out time);

                        data.Add(new ReportData
                        {
                            Date = date,
                            User1_M = row[30].ToString().Trim(),
                            US1_Prs1_M = row[31].ToString().Trim(),
                            Korrigiertefelder = korr,
                            Time = time
                        });
                    }

                    int month = dateTimePicker4.Value.Month;
                    int year = dateTimePicker4.Value.Year;

                    var report = data
                    .Where(x => !string.IsNullOrWhiteSpace(x.User1_M)
                             && x.US1_Prs1_M == "NK2"
                             && x.Date.Year == year
                             && x.Date.Month == month)
                    .GroupBy(x => new
                    {
                        Month = x.Date.ToString("yyyy-MM"),
                        User = x.User1_M
                    })
                    .Select(g => new
                    {
                        Month = g.Key.Month,
                        User = g.Key.User,
                        TotalKorrigiertefelder = g.Sum(x => x.Korrigiertefelder),
                        TotalTime = g.Sum(x => x.Time)
                    })
                    .ToList();

                    if (report.Count == 0)
                    {
                        MessageBox.Show("NK2 data not found.");
                        return;
                    }

                    var wb = new XLWorkbook();
                    var ws = wb.Worksheets.Add("NK2 Report");

                    // Header
                    ws.Cell(1, 1).Value = "Month";
                    ws.Cell(1, 2).Value = "User";
                    ws.Cell(1, 3).Value = "Korrigiertefelder";
                    ws.Cell(1, 4).Value = "Time";
                    ws.Cell(1, 5).Value = "PerHours";

                    int row1 = 2;

                    foreach (var item in report)
                    {
                        ws.Cell(row1, 1).Value = item.Month;
                        ws.Cell(row1, 2).Value = item.User;
                        ws.Cell(row1, 3).Value = item.TotalKorrigiertefelder;
                        ws.Cell(row1, 4).Value = item.TotalTime;
                        double perHours = 0;

                        if (item.TotalTime != 0)
                        {
                            perHours = Convert.ToDouble(item.TotalKorrigiertefelder) / Convert.ToDouble(item.TotalTime);
                        }

                        ws.Cell(row1, 5).Value = perHours;

                        row1++;
                    }

                    ws.Columns().AdjustToContents();

                    string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                    string folderPath = Path.Combine(desktopPath, "NK2_Reports");

                    if (!Directory.Exists(folderPath))
                    {
                        Directory.CreateDirectory(folderPath);
                    }

                    string filePath = Path.Combine(folderPath, "NK2_Report_" + month + "_" + year + ".xlsx");

                    wb.SaveAs(filePath);

                    MessageBox.Show("NK2 Report created successfully...");
                }
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            data.Clear();

            string sourceFile = @"\\192.168.0.223\Production Share\GDP207_Feedback\Report.xls";

            using (var stream = new FileStream(
     sourceFile,
     FileMode.Open,
     FileAccess.Read,
     FileShare.ReadWrite | FileShare.Delete))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet();

                    DataTable dt = result.Tables[0];




                    data.Clear();

                    for (int i = 1; i < dt.Rows.Count; i++)
                    {
                        DataRow row = dt.Rows[i];

                        if (string.IsNullOrWhiteSpace(row[39].ToString()))
                            continue;

                        DateTime date;
                        DateTime.TryParse(row[1].ToString(), out date);

                        int korr = 0;
                        int.TryParse(row[49].ToString(), out korr);

                        double time = 0;
                        double.TryParse(row[50].ToString(), out time);

                        data.Add(new ReportData
                        {
                            Date = date,
                            User1_M = row[39].ToString().Trim(),
                            US1_Prs1_M = row[40].ToString().Trim(),
                            Korrigiertefelder = korr,
                            Time = time
                        });
                    }

                    int month = dateTimePicker5.Value.Month;
                    int year = dateTimePicker5.Value.Year;

                    var report = data
                    .Where(x => !string.IsNullOrWhiteSpace(x.User1_M)
                             && x.US1_Prs1_M == "PZN"
                             && x.Date.Year == year
                             && x.Date.Month == month)
                    .GroupBy(x => new
                    {
                        Month = x.Date.ToString("yyyy-MM"),
                        User = x.User1_M
                    })
                    .Select(g => new
                    {
                        Month = g.Key.Month,
                        User = g.Key.User,
                        TotalKorrigiertefelder = g.Sum(x => x.Korrigiertefelder),
                        TotalTime = g.Sum(x => x.Time)
                    })
                    .ToList();

                    if (report.Count == 0)
                    {
                        MessageBox.Show("PZN data not found.");
                        return;
                    }

                    var wb = new XLWorkbook();
                    var ws = wb.Worksheets.Add("PZN Report");

                    // Header
                    ws.Cell(1, 1).Value = "Month";
                    ws.Cell(1, 2).Value = "User";
                    ws.Cell(1, 3).Value = "Korrigiertefelder";
                    ws.Cell(1, 4).Value = "Time";
                    ws.Cell(1, 5).Value = "PerHours";

                    int row1 = 2;

                    foreach (var item in report)
                    {
                        ws.Cell(row1, 1).Value = item.Month;
                        ws.Cell(row1, 2).Value = item.User;
                        ws.Cell(row1, 3).Value = item.TotalKorrigiertefelder;
                        ws.Cell(row1, 4).Value = item.TotalTime;
                        double perHours = 0;

                        if (item.TotalTime != 0)
                        {
                            perHours = Convert.ToDouble(item.TotalKorrigiertefelder) / Convert.ToDouble(item.TotalTime);
                        }

                        ws.Cell(row1, 5).Value = perHours;

                        row1++;
                    }

                    ws.Columns().AdjustToContents();

                    string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                    string folderPath = Path.Combine(desktopPath, "PZN_Reports");

                    if (!Directory.Exists(folderPath))
                    {
                        Directory.CreateDirectory(folderPath);
                    }

                    string filePath = Path.Combine(folderPath, "PZN_Report_" + month + "_" + year + ".xlsx");

                    wb.SaveAs(filePath);

                    MessageBox.Show("PZN Report created successfully...");
                }
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
             label6.Text = DateTime.Now.ToString(" dddd   |    dd MMMM yyyy     |     hh:mm:ss tt");

            // Left se right movement
            label6.Left += 1;

            // Right side se bahar jaane par wapas left
            if (label6.Left > this.ClientSize.Width)
            { 
                label6.Left = -label6.Width;
            }

           
        }
    }
}
