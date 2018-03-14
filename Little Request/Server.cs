using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using System.Threading;

namespace CBRE_Request {
    public class Server {

        private string[] _validUsers = new string[] { "TSandford" };
        
        private bool Running;
        private RequestLog Logger;
        private DateTime previousWriteTime;

        public Server(RequestLog logger) {
            this.Logger = logger;
        }

        public void Start() {
            Running = true;
            int sleepPeriodSeconds = 60;
            
            previousWriteTime = Logger.GetFileModified();

            while (Running) {
                Thread.Sleep(sleepPeriodSeconds * 1000);

                if (Logger.GetFileModified() > previousWriteTime) {
                    // a log has been added
                    var tempRequest = Logger.Read();
                    if (tempRequest.User != string.Empty) {
                        StartNotification(Logger, tempRequest);
                        Logger.Encrypt();
                    }
                    previousWriteTime = Logger.GetFileModified();
                }

            }
        }

        private void StartNotification(RequestLog logger, Request request) {
            var notification = new NotificationPopup(logger, request);
            notification.StartPosition = FormStartPosition.Manual;
            notification.Location = new System.Drawing.Point(Screen.PrimaryScreen.WorkingArea.Width - notification.Width,
                                                            Screen.PrimaryScreen.WorkingArea.Height - notification.Height);

            notification.TopMost = true;
            notification.ShowDialog();
        }

        public void Stop() {
            this.Running = false;
        }
        
        public bool HasAccess(string userName) {
            foreach(var user in this._validUsers)
                if (user.ToLower() == userName.ToLower())
                    return true;
            return false;
        }

    }
}
