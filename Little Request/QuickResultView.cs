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
    public partial class QuickResultView : Form {
        public QuickResultView(List<Request> listOfRequests) {
            InitializeComponent();
            SetUpDGV();
            LoadDGV(listOfRequests);
        }

        public QuickResultView() {
            InitializeComponent();
            SetUpDGV();
        }

        private void LoadDGV(List<Request> requests) {
            foreach(var request in requests) {
                int newRow = this.dataGridView1.Rows.Add();
                try {
                    this.dataGridView1.Rows[newRow].Cells["colDate"].Value = DateTime.Parse(request.Date).ToString("dd MMM yyyy hh:mm tt");
                } catch {
                    this.dataGridView1.Rows[newRow].Cells["colDate"].Value = request.Date;
                }
                this.dataGridView1.Rows[newRow].Cells["colUser"].Value = request.User;
                this.dataGridView1.Rows[newRow].Cells["colCategory"].Value = request.Category;
                this.dataGridView1.Rows[newRow].Cells["colUrgency"].Value = request.Urgency;
                this.dataGridView1.Rows[newRow].Cells["colDescription"].Value = request.Description;
                this.dataGridView1.Rows[newRow].Cells["colResponse"].Value = request.Response;
                this.dataGridView1.Rows[newRow].Cells["colCompletion"].Value = request.CompletionStatus;
            }
        }

        private void SetUpDGV() {
            var dgv = this.dataGridView1;
            dgv.AllowUserToResizeRows = true;
            dgv.AllowUserToDeleteRows = false;
            dgv.AllowUserToAddRows = false;
            dgv.AllowUserToOrderColumns = false;
            dgv.AllowUserToResizeColumns = true;
            dgv.Enabled = false;
            
            var cellTemplate = new DataGridViewTextBoxCell();
            dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;

            var DateCol = new DataGridViewColumn(cellTemplate);
            DateCol.HeaderText = "Date";
            DateCol.Name = "colDate";
            DateCol.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            DateCol.FillWeight = 10;

            var Usercol = new DataGridViewColumn(cellTemplate);
            Usercol.HeaderText = "User";
            Usercol.Name = "colUser";
            Usercol.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            Usercol.FillWeight = 10;

            var Category = new DataGridViewColumn(cellTemplate);
            Category.HeaderText = "Category";
            Category.Name = "colCategory";
            Category.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            Category.FillWeight = 15;

            var Urgency = new DataGridViewColumn(cellTemplate);
            Urgency.HeaderText = "Urgency";
            Urgency.Name = "colUrgency";
            Urgency.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            Urgency.FillWeight = 15;

            var Description = new DataGridViewColumn(cellTemplate);
            Description.HeaderText = "Description";
            Description.Name = "colDescription";
            Description.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            Description.FillWeight = 30;

            var Response = new DataGridViewColumn(cellTemplate);
            Response.HeaderText = "Response";
            Response.Name = "colResponse";
            Response.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            Response.FillWeight = 30;

            var Completion = new DataGridViewColumn(cellTemplate);
            Completion.HeaderText = "Completion";
            Completion.Name = "colCompletion";
            Completion.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            Completion.FillWeight = 10;

            dgv.Columns.Add(DateCol);
            dgv.Columns.Add(Usercol);
            dgv.Columns.Add(Category);
            dgv.Columns.Add(Urgency);
            dgv.Columns.Add(Description);
            dgv.Columns.Add(Response);
            dgv.Columns.Add(Completion);
                       

        }
    }
}
