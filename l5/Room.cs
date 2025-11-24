using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace l5
{
    public class Room
    {
        public int RoomId { get; set; }
        public int Floor { get; set; }
        public int Capacity { get; set; }
        public decimal PricePerDay { get; set; }
        public int Category { get; set; }


        public Room(int roomId, int floor, int capacity, decimal pricePerDay, int category)
        {
            RoomId = roomId;
            Floor = floor;
            Capacity = capacity;
            PricePerDay = pricePerDay;
            Category = category;
        }

        public override string ToString()
        {
            return $"№{RoomId}, этаж {Floor}, мест {Capacity}, {PricePerDay} р. /сутки, категория: {Category}";
        }

    }
}
