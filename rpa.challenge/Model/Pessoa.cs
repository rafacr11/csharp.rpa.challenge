using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace rpa.challenge.Model
{
    class Pessoa
    {      
        public String firstname { get; set; }
        public String lastname { get; set; }
        public String companyname { get; set; }
        public String phonenumber { get; set; }
        public String email { get; set; }
        public String role { get; set; }
        public String adress { get; set; }
        public bool isProcessed { get; set;}

        public Pessoa()
        {

        }
        public Pessoa(string firstname, string lastname, string companyname, string phonenumber, string email, string role, string adress,bool isProcessed)
        {
            this.firstname = firstname;
            this.lastname = lastname;
            this.companyname = companyname;
            this.phonenumber = phonenumber;
            this.email = email;
            this.role = role;
            this.adress = adress;
            this.isProcessed = isProcessed;
        }
    }
}
