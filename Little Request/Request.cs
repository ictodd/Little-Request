using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CBRE_Request {
    public struct Request {
        public string Date;
        public string User;
        public string Category;
        public string Urgency;
        public string Description;
        public int ID;
        public string Response;
        public string CompletionStatus;

        public Request(string date, string user, string categroy, string urgency, string description) {
            this.ID = -1;
            this.Date = date;
            this.User = user;
            this.Category = categroy;
            this.Urgency = urgency;
            this.Description = description;

            this.Response = "";
            this.CompletionStatus = "";
        }

    }
}
