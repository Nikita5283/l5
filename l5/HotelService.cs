using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;

namespace l5
{
    public static class HotelService
    {
        

        // Данные в памяти
        public static List<Client> Clients = new List<Client>();
        public static List<Room> Rooms = new List<Room>();
        public static List<Booking> Bookings = new List<Booking>();

        public static void LoadFromExcel(string path)
        {
            using var fs = new FileStream(path, FileMode.Open, FileAccess.Read);

            IWorkbook wb = WorkbookFactory.Create(fs);

            ISheet sheetClients = wb.GetSheetAt(0);
            ISheet sheetBooking = wb.GetSheetAt(1);
            ISheet sheetRooms = wb.GetSheetAt(2);

            Clients.Clear();
            Rooms.Clear();
            Bookings.Clear();

            // клиенты
            for (int i = 1; i <= sheetClients.LastRowNum; i++)
            {
                var row = sheetClients.GetRow(i);
                if (row == null) continue;

                try
                {
                    int id = (int)row.GetCell(0).NumericCellValue;
                    string last = row.GetCell(1).StringCellValue;
                    string first = row.GetCell(2).StringCellValue;
                    string mid = row.GetCell(3).StringCellValue;
                    string addr = row.GetCell(4).StringCellValue;

                    Clients.Add(new Client(id, last, first, mid, addr));
                }
                catch { }
            }

            // бронирование
            for (int i = 1; i <= sheetBooking.LastRowNum; i++)
            {
                var row = sheetBooking.GetRow(i);
                if (row == null) continue;

                try
                {
                    int bid = (int)row.GetCell(0).NumericCellValue;
                    int cid = (int)row.GetCell(1).NumericCellValue;
                    int rid = (int)row.GetCell(2).NumericCellValue;

                    DateTime book = GetDateFromCell(row.GetCell(3));
                    DateTime checkin = GetDateFromCell(row.GetCell(4));
                    DateTime checkout = GetDateFromCell(row.GetCell(5));


                    Bookings.Add(new Booking(bid, cid, rid, book, checkin, checkout));
                }
                catch { }
            }

            // номера
            for (int i = 1; i <= sheetRooms.LastRowNum; i++)
            {
                var row = sheetRooms.GetRow(i);
                if (row == null) continue;

                try
                {
                    int id = (int)row.GetCell(0).NumericCellValue;
                    int floor = (int)row.GetCell(1).NumericCellValue;
                    int cap = (int)row.GetCell(2).NumericCellValue;
                    decimal price = (decimal)row.GetCell(3).NumericCellValue;
                    int cat = (int)row.GetCell(4).NumericCellValue;

                    Rooms.Add(new Room(id, floor, cap, price, cat));
                }
                catch { }
            }
        }


        private static DateTime GetDateFromCell(ICell cell)
        {
            if (cell == null)
                throw new Exception("Cell is null");

            // Число (в Excel дата хранится как число)
            if (cell.CellType == CellType.Numeric)
            {
                if (DateUtil.IsCellDateFormatted(cell))
                    return cell.DateCellValue ?? throw new Exception("Numeric cell contains null date");

                throw new Exception("Numeric cell is not a date");
            }

            // Строка
            if (cell.CellType == CellType.String)
            {
                string s = cell.StringCellValue?.Trim();

                if (DateTime.TryParse(s, out var dt))
                    return dt;

                throw new Exception("Cannot parse string date: " + s);
            }

            // Формула
            if (cell.CellType == CellType.Formula)
            {
                if (cell.CachedFormulaResultType == CellType.Numeric)
                {
                    if (DateUtil.IsCellDateFormatted(cell))
                        return cell.DateCellValue ?? throw new Exception("Formula numeric date is null");

                    throw new Exception("Formula numeric is not a date");
                }

                if (cell.CachedFormulaResultType == CellType.String)
                {
                    string s = cell.StringCellValue?.Trim();

                    if (DateTime.TryParse(s, out var dt))
                        return dt;

                    throw new Exception("Cannot parse formula string date: " + s);
                }

                throw new Exception("Unsupported formula result type");
            }

            throw new Exception("Unsupported cell type: " + cell.CellType);
        }





        public static void SaveToExcel(string path)
        {
            IWorkbook wb;

            // Если файл существует — читаем и перезаписываем,
            // если нет — создаём новый .xlsx
            if (File.Exists(path))
            {
                using var fs = new FileStream(path, FileMode.Open, FileAccess.Read);
                wb = WorkbookFactory.Create(fs);
            }
            else
            {
                wb = new XSSFWorkbook();
            }

            // Удаляем старые листы, если были
            for (int i = wb.NumberOfSheets - 1; i >= 0; i--)
                wb.RemoveSheetAt(i);

            // Создаём чистые листы
            ISheet sheetClients = wb.CreateSheet("Клиенты");
            ISheet sheetBooking = wb.CreateSheet("Бронирование");
            ISheet sheetRooms = wb.CreateSheet("Номера");

            // клиенты
            WriteHeader(sheetClients, new[] { "Код клиента", "Фамилия", "Имя", "Отчество", "Место жительства" });

            int r = 1;
            foreach (var c in Clients)
            {
                IRow row = sheetClients.CreateRow(r++);
                row.CreateCell(0).SetCellValue(c.ClientId);
                row.CreateCell(1).SetCellValue(c.LastName);
                row.CreateCell(2).SetCellValue(c.FirstName);
                row.CreateCell(3).SetCellValue(c.Surname);
                row.CreateCell(4).SetCellValue(c.Address);
            }

            // бронирование
            WriteHeader(sheetBooking, new[] {
                "Код бронирования", "Код клиента", "Код номера",
                "Дата бронирования", "Дата заезда", "Дата выезда"
            });

            r = 1;
            foreach (var b in Bookings)
            {
                IRow row = sheetBooking.CreateRow(r++);
                row.CreateCell(0).SetCellValue(b.BookingId);
                row.CreateCell(1).SetCellValue(b.ClientId);
                row.CreateCell(2).SetCellValue(b.RoomId);
                row.CreateCell(3).SetCellValue(b.BookingDate.ToString("yyyy-MM-dd"));
                row.CreateCell(4).SetCellValue(b.CheckIn.ToString("yyyy-MM-dd"));
                row.CreateCell(5).SetCellValue(b.CheckOut.ToString("yyyy-MM-dd"));
            }

            // номера
            WriteHeader(sheetRooms, new[] {
                "Код номера", "Этаж", "Число мест",
                "Стоимость проживания", "Категория"
            });

            r = 1;
            foreach (var room in Rooms)
            {
                IRow row = sheetRooms.CreateRow(r++);
                row.CreateCell(0).SetCellValue(room.RoomId);
                row.CreateCell(1).SetCellValue(room.Floor);
                row.CreateCell(2).SetCellValue(room.Capacity);
                row.CreateCell(3).SetCellValue((double)room.PricePerDay);
                row.CreateCell(4).SetCellValue(room.Category);
            }

            // Сохраняем
            using var fsOut = new FileStream(path, FileMode.Create, FileAccess.Write);
            wb.Write(fsOut);
        }


        private static void WriteHeader(ISheet sheet, string[] headers)
        {
            IRow row = sheet.CreateRow(0);
            for (int i = 0; i < headers.Length; i++)
            {
                row.CreateCell(i).SetCellValue(headers[i]);
                sheet.AutoSizeColumn(i);
            }
        }

        // Просмотр всех таблиц (возвращает перечисления)
        public static IEnumerable<Client> ViewClients()
        {
            var q =
                from c in Clients
                orderby c.ClientId
                select c;

            return q;
        }

        public static IEnumerable<Room> ViewRooms()
        {
            var q =
                from r in Rooms
                orderby r.RoomId
                select r;

            return q;
        }

        public static IEnumerable<Booking> ViewBookings()
        {
            var q =
                from b in Bookings
                orderby b.BookingId
                select b;

            return q;
        }

        // Добавление элементов
        public static void AddClient(Client c)
        {
            if (Clients.Any(x => x.ClientId == c.ClientId)) // есть ли клиент, у которого ID совпадает с ID добавляемого?
                throw new InvalidOperationException("Клиент с таким ID уже есть.");
            Clients.Add(c);
        }

        public static void AddRoom(Room r)
        {
            if (Rooms.Any(x => x.RoomId == r.RoomId))
                throw new InvalidOperationException("Номер с таким ID уже есть.");
            Rooms.Add(r);
        }

        public static void AddBooking(Booking b)
        {
            if (Bookings.Any(x => x.BookingId == b.BookingId))
                throw new InvalidOperationException("Бронь с таким ID уже есть.");
            // Проверим, что клиент и номер существуют
            if (!Clients.Any(c => c.ClientId == b.ClientId))
                throw new InvalidOperationException("Клиент не найден.");
            if (!Rooms.Any(r => r.RoomId == b.RoomId))
                throw new InvalidOperationException("Номер не найден.");
            if (b.CheckIn >= b.CheckOut)
                throw new InvalidOperationException("Дата заезда должна быть раньше даты выезда.");
            Bookings.Add(b);
        }

        // Удаление по ключу
        public static bool DeleteClientById(int clientId)
        {
            var q =
                from c in Clients
                where c.ClientId == clientId
                select c;
            var toRemove = q.FirstOrDefault();
            if (toRemove != null)
            {
                // также удалим бронь этого клиента (или можно запретить удаление при наличии броней)
                Bookings.RemoveAll(b => b.ClientId == clientId);
                Clients.Remove(toRemove);
                return true;
            }
            return false;
        }

        public static bool DeleteRoomById(int roomId)
        {
            var q =
                from r in Rooms
                where r.RoomId == roomId
                select r;
            var toRemove = q.FirstOrDefault();
            if (toRemove != null)
            {
                // удалим связанные брони
                Bookings.RemoveAll(b => b.RoomId == roomId);
                Rooms.Remove(toRemove);
                return true;
            }
            return false;
        }

        public static bool DeleteBookingById(int bookingId)
        {
            var q =
                from b in Bookings
                where b.BookingId == bookingId
                select b;
            var toRemove = q.FirstOrDefault();
            if (toRemove != null)
            {
                Bookings.Remove(toRemove);
                return true;
            }
            return false;
        }

        // Запросы
        // - 1 запрос (1 таблица) -> возвращает перечень
        // - 1 запрос (2 таблицы) -> возвращает одно значение
        // - 2 запроса (3 таблицы) -> 1 перечень, 1 одно значение

        // список номеров заданной категории (перечень)
        public static IEnumerable<Room> GetRoomsByCategory(int category)
        {
            var q =
                from r in Rooms
                where r.Category == category
                select r;
            return q;
        }

        // число бронирований указанного клиента (одно значение)
        public static int GetBookingsCountByClientId(int clientId)
        {
            var q =
                from b in Bookings
                join c in Clients on b.ClientId equals c.ClientId
                where c.ClientId == clientId
                select b;
            return q.Count();
        }

        // перечень броней с именем клиента и номером комнаты (перечень)
        public static IEnumerable<object> GetBookingsDetailed()
        {
            var q =
                from b in Bookings
                join c in Clients on b.ClientId equals c.ClientId
                join r in Rooms on b.RoomId equals r.RoomId
                orderby b.CheckIn
                select new
                {
                    b.BookingId,
                    ClientName = c.FirstName,
                    RoomNumber = r.RoomId,
                    r.Category,
                    b.CheckIn,
                    b.CheckOut
                };
            return q;
        }

        // сумма предполагаемой выручки за все брони (одно значение)
        public static decimal GetTotalRevenue()
        {
            var q =
                from b in Bookings
                join r in Rooms on b.RoomId equals r.RoomId
                select new
                {
                    Days = (b.CheckOut - b.CheckIn).Days,
                    Price = r.PricePerDay
                };

            decimal total = 0m;
            foreach (var x in q)
            {
                int days = Math.Max(0, x.Days);
                total += x.Price * days;
            }
            return total;
        }
    }
}
