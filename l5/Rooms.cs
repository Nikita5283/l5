using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace l5
{
    internal class Rooms
    {
        private string code
        {
            get { return code; }
            set { code = value; }
        }
        private string floor
        {
            get { return floor; }
            set { floor = value; }
        }
        private string num_seats
        {
            get { return num_seats; }
            set { num_seats = value; }
        }
        private string cost_living
        {
            get { return cost_living; }
            set { cost_living = value; }
        }
        private string category
        {
            get { return category; }
            set { category = value; }
        }

        public Rooms(string code, string floor, string num_seats, string cost_living, string category)
        {
            this.code = code;
            this.floor = floor;
            this.num_seats = num_seats;
            this.cost_living = cost_living;
            this.category = category;
        }


    }
}
