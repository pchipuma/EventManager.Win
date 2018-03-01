using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace EventManager
{
    struct Person
    { 
        public string FirstName;
        public string Surname;
        public string Church;
        public string EmailAddress;
        public string CellPhoneNumber;
        public string WhatsappNo;
        public string WorkNumber;
        public string Program;
        public string Club;
        public string ReferenceNo;
    }
    public partial class frmEventManager : Form
    {
        string
            MasterFile = "C:\\Users\\us\\Downloads\\Rays 2018\\Master\\rays2018registration-master-report.xlsx",
            DownloadedFile = "C:\\Users\\us\\Downloads\\Rays 2018\\rays2018registration-report.xlsx",
            tempFile = "C:\\Users\\us\\Downloads\\Rays 2018\\temporary.xlsx";

        //string
        //    MasterFile = "J:\\Users\\Chipuma.Percy\\Downloads\\Rays 2018\\Master\\rays2018registration-master-report.xlsx",
        //    DownloadedFile = "J:\\Users\\Chipuma.Percy\\Downloads\\Rays 2018\\rays2018registration-report.xlsx",
        //    tempFile = "J:\\Users\\Chipuma.Percy\\Downloads\\Rays 2018\\temporary.xlsx";


        /// <summary>
        /// Initializes components
        /// </summary>
        public frmEventManager()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Sends out an email
        /// </summary>
        /// <param name="ReciverMail"></param>
        /// <returns></returns>
        private string SendEmail(Person recipient)
        {
            MailMessage msg = new MailMessage();
            msg.From = new MailAddress("wrays2018@gmail.com");
            msg.To.Add(recipient.EmailAddress);
            //msg.CC.Add("yakog@cc.adventist.org");
            //msg.CC.Add("carolusni@cc.adventist.org");
            //msg.Bcc.Add("magadzire.gary@gmail.com");
            msg.Bcc.Add("pchipuma@gmail.com");
            msg.Subject = "RAYS NOTIFICATION " + DateTime.Now.ToString();
            msg.Body = PrepareMessage(recipient);
            SmtpClient client = new SmtpClient();
            client.UseDefaultCredentials = true;
            client.Credentials = new NetworkCredential("wrays2018@gmail.com", "Log1n0I*");
            client.Host = "smtp.gmail.com";
            client.Port = 587;
            client.EnableSsl = true;
            client.DeliveryMethod = SmtpDeliveryMethod.Network;

            client.Timeout = 60000;
            try
            {
                client.Send(msg);
                return "Mail has been successfully sent!";
            }
            catch (Exception ex)
            {
                return "Fail Has error" + ex.Message;
            }
            finally
            {
                msg.Dispose();
            }
        }
        private void btnNotify_Click(object sender, EventArgs e)
        {
            try
            {
                //Open master file and get existing Ids.
                Excel.Application xlApp;
                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                Excel.Range range;
                int rCnt;
                int cCnt;
                int rowCount = 0;
                int columnCount = 0;
                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Open(MasterFile, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                range = xlWorkSheet.UsedRange;
                rowCount= range.Rows.Count;
                columnCount = range.Columns.Count;
                string[] registered = new string[rowCount-1];
                cCnt = 1;
                for (rCnt = 2; rCnt <= rowCount; rCnt++)
                {
                    registered[rCnt - 2] = (string)(range.Cells[rCnt, cCnt] as Excel.Range).Value2;
                }

                Excel.Application App = new Excel.Application();
                Excel.Workbook NewxlWorkBook = App.Workbooks.Open(DownloadedFile, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Excel.Worksheet NewxlWorkSheet = (Excel.Worksheet)NewxlWorkBook.Worksheets.get_Item(1);
                Excel.Range rng = NewxlWorkSheet.UsedRange;
                int rCount = rng.Rows.Count;
                int cCount = rng.Columns.Count;

                string[] newEntries = new string[rCount - 1];
                string Item = string.Empty;
                int c = 1;
                int startRow = rowCount + 2;

                uiProgressBar1.Minimum = 2;
                uiProgressBar1.Maximum = rCount;

                for (int r = 2; r <= rCount; r++)
                {
                    uiProgressBar1.Value = r;
                    Item = (string)(rng.Cells[r, c] as Excel.Range).Value2;

                    if (!string.IsNullOrEmpty(Item)&& !registered.Contains(Item))
                    {
                        Person person = new Person();
                        for (int i = 1; i < columnCount-1; i++)
                        {
                            xlWorkSheet.Cells[startRow, i] = (string)(rng.Cells[r, i] as Excel.Range).Value2;
                            xlWorkSheet.Cells[startRow, 14] = "YES";
                            switch (i)
                            { 
                                case 2:
                                    person.FirstName = (string)(rng.Cells[r, i] as Excel.Range).Value2;
                                    break;
                                case 3:
                                    person.Surname = (string)(rng.Cells[r, i] as Excel.Range).Value2;
                                    break;
                                case 4:
                                    person.EmailAddress = (string)(rng.Cells[r, i] as Excel.Range).Value2;
                                    break;
                                case 5:
                                    person.Church = (string)(rng.Cells[r, i] as Excel.Range).Value2;
                                    break;
                                case 6:
                                    person.CellPhoneNumber = (string)(rng.Cells[r, i] as Excel.Range).Value2;
                                    break;
                                case 7:
                                    person.WhatsappNo = (string)(rng.Cells[r, i] as Excel.Range).Value2;
                                    break;
                                case 8:
                                    person.WorkNumber = (string)(rng.Cells[r, i] as Excel.Range).Value2;
                                    break;
                                case 9:
                                    person.Program = (string)(rng.Cells[r, i] as Excel.Range).Value2;
                                    break;
                                case 10:
                                    person.Club = (string)(rng.Cells[r, i] as Excel.Range).Value2;
                                    break;
                                case 13:
                                    person.ReferenceNo = (string)(rng.Cells[r, i] as Excel.Range).Value2;
                                    break;
                            }

                        }
                        SendEmail(person);
                        startRow += 1;
                    }
                }
                NewxlWorkBook.Close(true, null, null);
                App.Quit();
                Marshal.ReleaseComObject(NewxlWorkSheet);
                Marshal.ReleaseComObject(NewxlWorkBook);
                Marshal.ReleaseComObject(App);

                xlWorkBook.Close(true, tempFile, Type.Missing);
                xlApp.Quit();
                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);

                if (File.Exists(tempFile))
                {
                    if (File.Exists(MasterFile))
                    {
                        File.Delete(MasterFile);
                    }
                    File.Move(tempFile, MasterFile);
                }

                if (File.Exists(DownloadedFile))
                {
                    string 
                        filename = Path.GetFileName(DownloadedFile),
                        folder = Path.GetDirectoryName(DownloadedFile),
                        destination = folder + "\\Processed\\" + GetTimestamp() + filename;
                    File.Move(DownloadedFile, destination);
  
                }

                MessageBox.Show("Notifications completed");
            }
            catch (Exception ex)
            {

                MessageBox.Show(string.Format("Please make sure the new file is inside the Rays 2018 folder. \n {0}", ex));
            }
        }

        private string PrepareMessage(Person recipient)
        {
            string EmailBody = string.Format("Good day {0} \n\nThank you for registering for RAYS 2018. Herewith are your registered details: \n\nName: \t\t{1} {2} \nChurch: \t\t{3} \nE-mail: \t\t{4} \nCellphone No.: \t{5} \nWhatsApp No.: \t{6} \nOther No.: \t{7} \nProgram: \t\t{8} - {9} \n\nIf you confirm that the details supplied are correct, please proceed to make a payment of R180 into the following bank account: \n\n Bank: \t\tStandard Bank \nAccount Name: \tWestern Region Adventist Youth Ministry/ WRAYM \nAccount No.: \t078 320 739 \nReference: \t{10} \n\nNB: The reference number is unique for your registration. If you registered for multiple attendees your co-attendees will get e-mails with the same reference number (please confirm with them). If the reference is the same, you can make a single payment using that same reference (amounting to R180 x attendees with the same reference number), otherwise, proceed to make a payment of R180 only :) \n\nPlease send proof of payment to yakog@cc.adventist.org and/or WhatsApp XYZ \n\nThank you and hoping to seeing you at RAYS. \n\nYours in Christ, \nThe WRAYM Team", recipient.FirstName, recipient.FirstName, recipient.Surname, recipient.Church, recipient.EmailAddress, recipient.CellPhoneNumber, recipient.WhatsappNo, recipient.WorkNumber, recipient.Program.ToUpper(), recipient.Club.ToUpper(), recipient.ReferenceNo.ToUpper());
            return EmailBody;
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            dlgOpenFile.ShowDialog();
            DownloadedFile = dlgOpenFile.FileName;
            txtNewFile.Text = DownloadedFile;
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            //SendEmail("pchipuma@hotmail.com");
        }

        private void frmEventManager_Load(object sender, EventArgs e)
        {
            txtNewFile.Text = DownloadedFile;
        }

        private string GetTimestamp()
        {
            DateTime timestamp = DateTime.Now;
            string 
                yy = timestamp.Year.ToString(),
                mm = timestamp.Month < 10 ? "0" + timestamp.Month.ToString() : timestamp.Month.ToString(),
                dd = timestamp.Day < 10 ? "0" + timestamp.Day.ToString() : timestamp.Day.ToString(),
                hh = timestamp.Hour < 10 ? "0" + timestamp.Hour.ToString() : timestamp.Hour.ToString(),
                mn = timestamp.Minute < 10 ? "0" + timestamp.Minute.ToString() : timestamp.Minute.ToString();
            return yy + mm + dd + "–" + hh + "h" + mn;
        }
    }
}
