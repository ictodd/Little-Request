using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.IO;
using System.Collections.Generic;

namespace CBRE_Request {
    public class RequestLog {

        private static string REQUEST_RECORDS_FILEPATH_DECRYPTED = @"\\nzaklfnp03\Departments\Valuations\TODDS APP STORE\Programs\CBRERequestLog.csv";

        private static string REQUEST_RECORDS_FILEPATH_ENCRYPTED = @"\\nzaklfnp03\Departments\Valuations\TODDS APP STORE\Programs\CBRERequestLog.vas";


        public RequestLog() { }

        public List<Request> ReadAll() {
            Decrypt();

            List<Request> result = new List<Request>();

            if (IsFileLocked(REQUEST_RECORDS_FILEPATH_DECRYPTED)) return result;

            Excel.Application xl = null;
            Excel.Workbook wb = null;
            Excel.Workbooks books = null;
            Excel.Worksheet ws = null;


            try {
                xl = new Excel.Application();
                books = xl.Workbooks;
                wb = books.Open(REQUEST_RECORDS_FILEPATH_DECRYPTED);
                ws = wb.Worksheets[1];
                int lr = ws.Range["A" + ws.Rows.Count].End[Excel.XlDirection.xlUp].Row;

                for(int i = 2; i <= lr; i++) {
                    double temp;
                    Request tempResult = new Request();
                    try {
                        // reading dates from excel is annoying...
                        temp = double.Parse(ws.Range["A" + i].Value2.ToString());

                        tempResult.Date = DateTime.FromOADate(temp).ToString("d MMM yyyy hh:mm tt");
                    } catch {
                        tempResult.Date = "unknown";
                    }
                    tempResult.User = ws.Range["B" + i].Value2;
                    tempResult.Category = ws.Range["C" + i].Value2;
                    tempResult.Urgency = ws.Range["D" + i].Value2;
                    tempResult.Description = ws.Range["E" + i].Value2;
                    tempResult.Response = ws.Range["F" + i].Value2;
                    tempResult.CompletionStatus = ws.Range["G" + i].Value2;
                    tempResult.ID = i;
                    result.Add(tempResult);
                }

                wb.Save();
                wb.Close();
                xl.Quit();
            } catch (Exception ex) {
                MessageBox.Show($"Error reading from log.\nError: {ex.Message}");
            } finally {
                if (books != null) Marshal.ReleaseComObject(books);
                if (wb != null) Marshal.ReleaseComObject(wb);
                if (ws != null) Marshal.ReleaseComObject(ws);
                if (xl != null) Marshal.ReleaseComObject(xl);
            }
            Encrypt();
            return result;
        }

        public Request Read() {
            
            Decrypt();

            Request result = new Request();

            if (IsFileLocked(REQUEST_RECORDS_FILEPATH_DECRYPTED)) return result;

            Excel.Application xl = null;
            Excel.Workbook wb = null;
            Excel.Workbooks books = null;
            Excel.Worksheet ws = null;
            

            try {
                xl = new Excel.Application();
                books = xl.Workbooks;
                wb = books.Open(REQUEST_RECORDS_FILEPATH_DECRYPTED);
                ws = wb.Worksheets[1];
                int lr = ws.Range["A" + ws.Rows.Count].End[Excel.XlDirection.xlUp].Row;

                double temp;
                try {
                    // reading dates from excel is annoying...
                    temp = double.Parse(ws.Range["A" + lr].Value2.ToString());

                    result.Date = DateTime.FromOADate(temp).ToString("d MMM yyyy hh:mm tt");
                } catch {
                    result.Date = "unknown";
                }
                result.User = ws.Range["B" + lr].Value2;
                result.Category = ws.Range["C" + lr].Value2;
                result.Urgency = ws.Range["D" + lr].Value2;
                result.Description = ws.Range["E" + lr].Value2;
                result.ID = lr;
                wb.Save();
                wb.Close();
                xl.Quit();
            } catch (Exception ex) {
                MessageBox.Show($"Error reading from log.\nError: {ex.Message}");
            } finally {
                if (books != null) Marshal.ReleaseComObject(books);
                if (wb != null) Marshal.ReleaseComObject(wb);
                if (ws != null) Marshal.ReleaseComObject(ws);
                if (xl != null) Marshal.ReleaseComObject(xl);
            }
            Encrypt();
            return result;
        }

        public bool WriteResponse(int row, string responseMessage, string completionStatus) {
            if(row < 2) 
                return false;
            
            Decrypt();

            if (IsFileLocked(REQUEST_RECORDS_FILEPATH_DECRYPTED)) {
                MessageBox.Show("Could not log. Please try again soon.");
                return false;
            }

            Excel.Application xl = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;
            Excel.Workbooks books = null;

            bool result;
            try {
                xl = new Excel.Application();
                books = xl.Workbooks;
                wb = xl.Workbooks.Open(REQUEST_RECORDS_FILEPATH_DECRYPTED);
                ws = wb.Worksheets[1];
                ws.Range["F" + row].Value2 = responseMessage;
                ws.Range["G" + row].Value2 = completionStatus;
                wb.Save();
                wb.Close();
                xl.Quit();
                result = true;
            } catch (Exception ex) {
                MessageBox.Show($"Error writing to log.\nError: {ex.Message}");
                result = false;
            } finally {
                if (books != null) Marshal.ReleaseComObject(books);
                if (wb != null) Marshal.ReleaseComObject(wb);
                if (ws != null) Marshal.ReleaseComObject(ws);
                if (xl != null) Marshal.ReleaseComObject(xl);
            }
            Encrypt();
            return result;
        }

        public DateTime GetFileModified() {
            DateTime result = new DateTime();
            FileInfo fi;
            Decrypt();
            if (!IsFileLocked(REQUEST_RECORDS_FILEPATH_DECRYPTED)) {
                if (File.Exists(REQUEST_RECORDS_FILEPATH_DECRYPTED)) {
                    fi = new FileInfo(REQUEST_RECORDS_FILEPATH_DECRYPTED);
                    result = fi.LastWriteTime;
                }
            }
            fi = null;
            Encrypt();
            return result;
        }

        public bool Write (Request request){
            
            Decrypt();

            if (IsFileLocked(REQUEST_RECORDS_FILEPATH_DECRYPTED)) {
                MessageBox.Show("Could not log. Please try again soon.");
                return false;
            }

            Excel.Application xl = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;
            Excel.Workbooks books = null;

            bool result;
            try {
                xl = new Excel.Application();
                books = xl.Workbooks;
                wb = xl.Workbooks.Open(REQUEST_RECORDS_FILEPATH_DECRYPTED);
                ws = wb.Worksheets[1];
                int lr = ws.Range["A" + ws.Rows.Count].End[Excel.XlDirection.xlUp].Row + 1;
                ws.Range["A" + lr].Value2 = request.Date;
                ws.Range["B" + lr].Value2 = request.User;
                ws.Range["C" + lr].Value2 = request.Category;
                ws.Range["D" + lr].Value2 = request.Urgency;
                ws.Range["E" + lr].Value2 = request.Description.Replace("\n"," ").Replace("\r"," ");
                wb.Save();
                wb.Close();
                xl.Quit();
                result = true;
            } catch (Exception ex){
                MessageBox.Show($"Someone is viewing the request log (most likely Todd). Please try again soon.\nError: {ex.Message}");
                result = false;
            } finally {
                if (books != null) Marshal.ReleaseComObject(books);
                if (wb != null) Marshal.ReleaseComObject(wb);
                if (ws != null) Marshal.ReleaseComObject(ws);
                if (xl != null) Marshal.ReleaseComObject(xl);
            }
            Encrypt();
            return result;
        }

        public void Encrypt() {
            // i use this term ironically
            //if (IsFileLocked(REQUEST_RECORDS_FILEPATH_DECRYPTED)) return;

            if (File.Exists(REQUEST_RECORDS_FILEPATH_DECRYPTED))
                File.Move(REQUEST_RECORDS_FILEPATH_DECRYPTED,
                    Path.ChangeExtension(REQUEST_RECORDS_FILEPATH_DECRYPTED, "vas"));
        }
                
        public void Decrypt() {
            // i use this term ironically
            //if (IsFileLocked(REQUEST_RECORDS_FILEPATH_ENCRYPTED)) return;

            if (File.Exists(REQUEST_RECORDS_FILEPATH_ENCRYPTED))
                File.Move(REQUEST_RECORDS_FILEPATH_ENCRYPTED,
                    Path.ChangeExtension(REQUEST_RECORDS_FILEPATH_ENCRYPTED, "csv"));
        }

        public void OpenRequestLog() {
            Decrypt();

            Excel.Application xl = null;
            Excel.Workbook wb = null;
            Excel.Workbooks books = null;
            Excel.Worksheet ws = null;


            try {
                xl = new Excel.Application();
                xl.Visible = true;
                books = xl.Workbooks;
                wb = books.Open(REQUEST_RECORDS_FILEPATH_DECRYPTED);
                wb.Activate();
            } catch (Exception ex) {
                MessageBox.Show($"Error reading from log.\nError: {ex.Message}");
            } finally {
                if (books != null) Marshal.ReleaseComObject(books);
                if (wb != null) Marshal.ReleaseComObject(wb);
                if (ws != null) Marshal.ReleaseComObject(ws);
                if (xl != null) Marshal.ReleaseComObject(xl);
            }
        }

        public bool IsFileLocked(string filePath) {
            FileStream stream = null;
            FileInfo file = new FileInfo(filePath);

            try {
                stream = file.Open(FileMode.Open, FileAccess.Read, FileShare.None);
            } catch (IOException) {
                //the file is unavailable because it is:
                //still being written to
                //or being processed by another thread
                //or does not exist (has already been processed)
                return true;
            } finally {
                if (stream != null)
                    stream.Close();
            }

            //file is not locked
            return false;
        }
    }
}
