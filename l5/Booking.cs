using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace l5
{
    public class Booking
    {
        public int BookingId { get; set; }
        public int ClientId { get; set; }
        public int RoomId { get; set; }
        public DateTime BookingDate { get; set; }
        public DateTime CheckIn { get; set; }
        public DateTime CheckOut { get; set; }


        public Booking(int bookingId, int clientId, int roomId, DateTime bookingDate, DateTime checkIn, DateTime checkOut)
        {
            BookingId = bookingId;
            ClientId = clientId;
            RoomId = roomId;
            BookingDate = bookingDate;
            CheckIn = checkIn;
            CheckOut = checkOut;
        }

        public override string ToString()
        {
            return $"Бронь #{BookingId}: Клиент {ClientId}, Номер {RoomId}, Бронь {BookingDate:dd.MM.yyyy}, Заезд {CheckIn:dd.MM.yyyy}, Выезд {CheckOut:dd.MM.yyyy}";
        }
    }
}
