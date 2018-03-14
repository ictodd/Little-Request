using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace CBRE_Request {
    public partial class MainForm : Form {

        private Server _server;

        private string[] BlockedWords = new string[] {"beer", "fuck", "suck", "shit", "drink",
            "balls", "anus", "arsehole", "penis", "pussy", "beverage", "coffee", "alcohol", "dick", "bastard",
        };

        public MainForm() {
            InitializeComponent();
            LoadComboBoxes();
            SetUpDisabledFields();
            UpdateStatusBar("Idle");
        }

        private void SetUpDisabledFields() {

            this.txtUsername.Enabled = false;
            this.txtDate.Enabled = false;

            this.txtUsername.Text = Environment.UserName;
            this.txtDate.Text = DateTime.Now.ToString("d MMM yyyy h:mm tt");
        }

        private void LoadComboBoxes() {
            this.cboRequestCategory.Items.Add("Forbury");
            this.cboRequestCategory.Items.Add("New Word Report");
            this.cboRequestCategory.Items.Add("Linking");
            this.cboRequestCategory.Items.Add("General Excel");
            this.cboRequestCategory.Items.Add("General Word");
            this.cboRequestCategory.Items.Add("Other");

            this.cboUrgency.Items.Add("High");
            this.cboUrgency.Items.Add("Medium");
            this.cboUrgency.Items.Add("Low");
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e) {
            var ans = MessageBox.Show("Are you sure you want to exit?", "Exit", MessageBoxButtons.YesNo);
            if (ans == DialogResult.Yes) Environment.Exit(0);
        }

        private void desktopShortcutToolStripMenuItem_Click(object sender, EventArgs e) {
            var ans = MessageBox.Show("Are you sure you want to add a desktop shortcut?", "Desktop Shortcut", MessageBoxButtons.YesNo);
            if (ans == DialogResult.Yes) {
                if (File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\CBRE Request.lnk")) {
                    MessageBox.Show("You already have a shortcut on your desktop to this application.\n\rPlease don't be greedy.");
                    return;
                }

                IWshRuntimeLibrary.IWshShell_Class wsh = new IWshRuntimeLibrary.IWshShell_Class();
                IWshRuntimeLibrary.IWshShortcut shortcut = wsh.CreateShortcut(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\CBRE Request.lnk") as IWshRuntimeLibrary.IWshShortcut;
                shortcut.Arguments = "";
                shortcut.TargetPath = System.Reflection.Assembly.GetEntryAssembly().Location;
                shortcut.WindowStyle = 1;
                shortcut.Description = "Shortcut to CBRE Request system";
                shortcut.Save();
                MessageBox.Show("Desktop shortcut made!");

            }

        }

        private void btnSendRequest_Click(object sender, EventArgs e) {
            if (!ValidEntry()) return;

            Request request = new Request() {
                Date = this.txtDate.Text,
                User = this.txtUsername.Text,
                Category = this.cboRequestCategory.Text,
                Urgency = this.cboUrgency.Text,
                Description = this.txtBriefDescription.Text
            };

            RequestLog logger = new RequestLog();
            UpdateStatusBar("Logging Request...");
            if (logger.Write(request)) {
                // request logging was successful
                UpdateStatusBar("Idle");
                var ans = MessageBox.Show("Result successfully logged. Would you like to log another?", "Logging Success", MessageBoxButtons.YesNo);
                if(ans == DialogResult.No) {
                    Application.Exit();
                }
            }
            UpdateStatusBar("Idle");

        }

        private bool ValidEntry() {
            if(this.cboRequestCategory.Text == "") {
                MessageBox.Show("Please select a request category.");
                this.cboRequestCategory.Focus();
                return false;
            }
            if(this.cboUrgency.Text == "") {
                MessageBox.Show("Please select a level of urgency.");
                this.cboUrgency.Focus();
                return false;
            }
            if(this.txtBriefDescription.Text == "") {
                MessageBox.Show("Please provide a brief description of the request.");
                this.txtBriefDescription.Focus();
                return false;
            }
            if (CheckUnprofessionalDescription(this.txtBriefDescription.Text)) {
                MessageBox.Show("Please try writing the description in a professional manner.");
                this.txtBriefDescription.Focus();
                return false;
            }
            return true;
        }

        private bool CheckUnprofessionalDescription(string briefDescription) {
            foreach(var blockedWord in this.BlockedWords) {
                if(briefDescription.ToLower().IndexOf(blockedWord) >= 0) {
                    return true;
                }
            }
            return false;
        }

        private void startToolStripMenuItem_Click(object sender, EventArgs e) {
            this._server = new Server(new RequestLog());
            if (!this._server.HasAccess(this.txtUsername.Text)) {
                MessageBox.Show("You do not have the rights to start a CBRE Request server.");
            }
            UpdateStatusBar("Server started...");
            var serverTask = Task.Run(() => this._server.Start());
            UpdateStatusBar("Server running...");
        }

        private void stopToolStripMenuItem_Click(object sender, EventArgs e) {
            if (Environment.UserName.ToLower() != "tsandford") {
                MessageBox.Show("You don't have permission to stop the server.");
                return;
            }

            if (this._server != null) {
                this._server.Stop();
                UpdateStatusBar("Idle");
                MessageBox.Show("Server stopped.");
                this._server = null;
            }
        }

        public void UpdateStatusBar(string msg) {
            this.StatusBar.Text = $"CBRE Request Status: {msg}";
        }

        private async void quickLogViewToolStripMenuItem_Click(object sender, EventArgs e) {
            if(Environment.UserName.ToLower() != "tsandford") {
                MessageBox.Show("You don't have permission to view the request log.");
                return;
            }

            RequestLog logger = new RequestLog();
            UpdateStatusBar("Loading Quick Log...");
            var requests = await Task.Run(()=> logger.ReadAll());
            var resultView = new QuickResultView(requests);
            UpdateStatusBar("Currently viewing Quick Log...");
            resultView.ShowDialog();
            UpdateStatusBar("Idle");
        }

        private async void normalLogViewToolStripMenuItem_Click(object sender, EventArgs e) {
            if (Environment.UserName.ToLower() != "tsandford") {
                MessageBox.Show("You don't have permission to view the request log.");
                return;
            }


            RequestLog logger = new RequestLog();
            UpdateStatusBar("Loading Log...");
            await Task.Run(() => logger.OpenRequestLog());
            UpdateStatusBar("Idle");

        }
    }
}
