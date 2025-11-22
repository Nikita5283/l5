using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace l5
{
    public class Client
    {
        public int ClientId { get; set; }
        public string LastName { get; set; }
        public string FirstName { get; set; }
        public string Surname { get; set; }
        public string Address { get; set; }

        public Client() { }

        public Client(int clientId, string lastName, string firstName, string surname, string address)
        {
            ClientId = clientId;
            LastName = lastName;
            FirstName = firstName;
            Surname = surname;
            Address = address;
        }

        public override string ToString()
        {
            return $"{ClientId}: {LastName} {FirstName} {Surname} ({Address})";
        }
    }
}
