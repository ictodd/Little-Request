using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CBRE_Request {
    public partial class NotificationPopup : Form {
        private Request _request;
        private RequestLog _logger;

        public NotificationPopup(RequestLog logger, Request request) {
            InitializeComponent();
            this._request = request;
            this._logger = logger;

            LoadRequest();
            this.cboCompleted.Items.Add("Yes");
            this.cboCompleted.Items.Add("No");
            this.cboCompleted.Text = "";
        }

        private void LoadRequest() {
            this.lblUsername.Text = this._request.User;
            this.lblTime.Text = this._request.Date;
            this.lblCategory.Text = this._request.Category;
            this.lblUrgency.Text = this._request.Urgency;
            this.txtDesc.Text = this._request.Description;
            this.lblID.Text = this._request.ID.ToString();
        }

        private void btnOk_Click(object sender, EventArgs e) {
            if(this.cboCompleted.Text == "") {
                MessageBox.Show("Please select a 'completed status'.");
                this.cboCompleted.Focus();
                return;
            }
            if(this.txtResponse.Text == "") {
                MessageBox.Show("Please select fill in a response.");
                this.txtResponse.Focus();
                return;
            }

            int row = int.Parse(this.lblID.Text);
            if(_logger.WriteResponse(row, this.txtResponse.Text, this.cboCompleted.Text)) {
                MessageBox.Show("Successfully updated log with response.");
            } else {
                MessageBox.Show("Could not update log. Try again soon.");
            }

        }
    }
}
