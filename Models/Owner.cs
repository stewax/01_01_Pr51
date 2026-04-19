using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Word_kazakov.Models
{
    public class Owner
    {
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string SureName { get; set; }
        public int NumberRoom { get; set; }

        public Owner(string FirstName, string LastName, string SureName, int dNumberRoom)
        {
            this.FirstName = FirstName;
            this.LastName = LastName;
            this.SureName = SureName;
            this.NumberRoom = dNumberRoom;
        }
    }
}
