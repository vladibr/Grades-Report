using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using System.Net.Mail;
using System.Net;
using System.IO;
using System.Threading;
using System.Security.Cryptography;
using System.Data;

namespace Grades_Report_new
{
    class Program
    {
        static void Main(string[] args)
        {
            string connectionString = "Data Source=DESKTOP-BVVQLS1\\SQLEXPRESS;Initial Catalog=Ruppin_Students;Integrated Security=True;Connect Timeout=30;Encrypt=False;TrustServerCertificate=False;ApplicationIntent=ReadWrite;MultiSubnetFailover=False";
         

            string username;
            string password;
            float average;


            string queryString = "DELETE FROM [courses]\n" +
                                 "SELECT * FROM students";

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                SqlCommand command = new SqlCommand(queryString, connection);
                connection.Open();
                
                SqlDataReader reader = command.ExecuteReader();

                while (reader.Read())
                {
                    var psi = new ProcessStartInfo();
                    psi.FileName = @"D:\Python\Python310\python.exe";

                    var script = @"D:\Salman_Api\Grades_Report_Query\Grades_Report_new\Grades_Report_new\PDFRequests2.py";

                    username = reader["username_id"].ToString().Trim();
                    password = reader["password"].ToString();
                    password = Decrypt(password, username);

                    psi.Arguments = string.Format("{0} {1} {2} {3}",
                                    script,
                                    username,
                                    password,
                                    username);

                    psi.UseShellExecute = false;
                    psi.CreateNoWindow = true;
                    psi.RedirectStandardOutput = true;
                    psi.RedirectStandardError = true;

                    var process = Process.Start(psi);
                    process.WaitForExit();

                    using (var conn = new SqlConnection(connectionString))
                    {
                        conn.Open();

                        var cmd = new SqlCommand($"EXEC insert_grades {username}", conn);
                        cmd.ExecuteReader();

                        conn.Close();
                    }

                    using (var conn = new SqlConnection(connectionString))
                    {
                        conn.Open();

                        var cmd = new SqlCommand($"user_average", conn);
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@user_id", username);
                        cmd.Parameters.Add("@RetVal", SqlDbType.Float).Direction = ParameterDirection.ReturnValue;

                        cmd.ExecuteNonQuery();
                        average = Convert.ToInt32(cmd.Parameters["@RetVal"].Value);

                        conn.Close();
                    }

                    try
                    {
                        MailMessage message = new MailMessage();
                        SmtpClient smtp = new SmtpClient();
                        message.From = new MailAddress("Chaikeren0@gmail.com");
                        message.To.Add(reader["email"].ToString());
                        message.Subject = "Grades Report from Yedion";
                        message.Body = $"Your average grade: {average}";
                        message.IsBodyHtml = true;
                        string FileName = $"D:\\Salman_Api\\Grades_Report_Query\\Reports\\{username}.xlsx";
                        message.Attachments.Add(new Attachment(FileName));
                        smtp.Port = 587;
                        smtp.Host = "smtp.gmail.com";
                        smtp.EnableSsl = true;
                        smtp.UseDefaultCredentials = false;
                        smtp.Credentials = new NetworkCredential("Chaikeren0@gmail.com", "Chaikeren3131@");
                        smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                        smtp.Send(message);
                        smtp.Dispose();
                    }
                    catch (Exception a)
                    {
                        Console.WriteLine(a.Message);
                    }
                }

                // Call Close when done reading.
                reader.Close();
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
            string[] files = Directory.GetFiles(@"D:\Salman_Api\Grades_Report_Query\Reports");
            foreach(string file in files)
            {
                File.Delete(file);
            }
        }
        public static string Encrypt(string encryptString, string encryptionKey)
        {
            byte[] clearBytes = Encoding.Unicode.GetBytes(encryptString);
            using (Aes encryptor = Aes.Create())
            {
                Rfc2898DeriveBytes pdb = new Rfc2898DeriveBytes(encryptionKey, new byte[] {
                0x49, 0x76, 0x61, 0x6e, 0x20, 0x4d, 0x65, 0x64, 0x76, 0x65, 0x64, 0x65, 0x76
                });
                encryptor.Key = pdb.GetBytes(32);
                encryptor.IV = pdb.GetBytes(16);
                using (MemoryStream ms = new MemoryStream())
                {
                    using (CryptoStream cs = new CryptoStream(ms, encryptor.CreateEncryptor(), CryptoStreamMode.Write))
                    {
                        cs.Write(clearBytes, 0, clearBytes.Length);
                        cs.Close();
                    }
                    encryptString = Convert.ToBase64String(ms.ToArray());
                }
            }
            return encryptString;
        }
        public static string Decrypt(string cipherText, string encryptionKey)
        {
            cipherText = cipherText.Replace(" ", "+");
            byte[] cipherBytes = Convert.FromBase64String(cipherText);
            using (Aes encryptor = Aes.Create())
            {
                Rfc2898DeriveBytes pdb = new Rfc2898DeriveBytes(encryptionKey, new byte[] {
                 0x49, 0x76, 0x61, 0x6e, 0x20, 0x4d, 0x65, 0x64, 0x76, 0x65, 0x64, 0x65, 0x76
                 });
                encryptor.Key = pdb.GetBytes(32);
                encryptor.IV = pdb.GetBytes(16);
                using (MemoryStream ms = new MemoryStream())
                {
                    using (CryptoStream cs = new CryptoStream(ms, encryptor.CreateDecryptor(), CryptoStreamMode.Write))
                    {
                        cs.Write(cipherBytes, 0, cipherBytes.Length);
                        cs.Close();
                    }
                    cipherText = Encoding.Unicode.GetString(ms.ToArray());
                }
            }
            return cipherText;
        }
    }
}
