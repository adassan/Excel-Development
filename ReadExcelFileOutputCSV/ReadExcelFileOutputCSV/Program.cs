using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.IO;
using System.Data;
using Lesnikowski.Client;
using Lesnikowski.Mail;
using Lesnikowski.Mail.Headers;

namespace ReadExcelFileOutputCSV
{
    public class Program
    {
        private const string _server = "mail.peakproducts.com";
        private const string _user = "adassan@peakproducts.com";
        private const string _password = "d@ss@n";

        public static void Main(string[] args)
        {
            getAttachmentsMailDLL();
            //string sourceFile, worksheetName, targetFile;
            //sourceFile = "source.xls"; 
            //worksheetName = "sheet1"; 
            //targetFile = "target.csv";
            //targetFile = "C:\\Users\\Ashu\\Documents\\target.csv";
            //convertExcelToCSV(sourceFile, worksheetName, targetFile);            
        }

        static void convertExcelToCSV(string sourceFile, string worksheetName, string targetFile)
        {
            string strFileName = "C:\\Users\\Ashu\\Documents\\source.xls";
            string strConn = @"Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + strFileName + ";" + "Extended Properties=" + "\"" + "Excel 12.0;HDR=YES;" + "\"";

            OleDbConnection conn = null;
            StreamWriter wrtr = null;
            OleDbCommand cmd = null;
            OleDbDataAdapter da = null;

            try
            {
                conn = new OleDbConnection(strConn);
                conn.Open();

                cmd = new OleDbCommand("SELECT * FROM [" + worksheetName + "$]", conn);
                cmd.CommandType = CommandType.Text;
                wrtr = new StreamWriter(targetFile);

                da = new OleDbDataAdapter(cmd);
                DataTable dt = new DataTable();

                da.Fill(dt);

                for (int x = 0; x < dt.Rows.Count; x++)
                {
                    string rowString = "";
                    for (int y = 0; y < dt.Columns.Count; y++)
                    {
                        rowString += "\"" + dt.Rows[x][y].ToString() + "\",";
                    }

                    wrtr.WriteLine(rowString);
                }

                Console.WriteLine();
                Console.WriteLine("Done! Your " + sourceFile + " has been converted into " + targetFile + ".");
                Console.WriteLine();

            }

            catch (Exception exc)
            {
                Console.WriteLine(exc.ToString());
                Console.ReadLine();
            }
            finally
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();

                conn.Dispose();
                cmd.Dispose();
                da.Dispose();
                wrtr.Close();
                wrtr.Dispose();

            }

        }

        static void downloadAttachments()
        {
            Chilkat.MailMan mail = new Chilkat.MailMan();

        }

        static void getAttachmentsMailDLL()
        {
            using (Pop3 pop3 = new Pop3())
            {
                pop3.Connect(_server);                      // Use overloads or ConnectSSL if you need to specify different port or SSL.
                pop3.Login(_user, _password);               // You can also use: LoginAPOP, LoginPLAIN, LoginCRAM, LoginDIGEST methods,
                // or use UseBestLogin method if you want Mail.dll to choose for you.

                List<string> uidList = pop3.GetAll();       // Get unique-ids of all messages.

                foreach (string uid in uidList)
                {
                    IMail email = new MailBuilder().CreateFromEml(  // Download and parse each message.
                        pop3.GetMessageByUID(uid));

                    if ((email.Date <= DateTime.Now) && (email.Date >= DateTime.Today.AddDays(-1)))
                    {
                        ProcessMessage(email);                          // Display email data, save attachments.
                    }
                }

                pop3.Close();
            }
        }

        private static void ProcessMessage(IMail email)
        {
            Console.WriteLine("Attachments: ");
            foreach (MimeData attachment in email.Attachments)
            {
                string fileName = attachment.FileName;
                int fileExtPos = fileName.LastIndexOf(".");
                if (fileExtPos >= 0)
                    fileName = fileName.Substring(0, fileExtPos);

                Console.WriteLine(fileName);
                attachment.Save(@"c:\" + fileName);
            }
        }
    }
}
