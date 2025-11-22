using Aspose.Cells;
using l5;
using System.Data;
using System.Globalization;

namespace l5
{
    internal class Program
    {
        private static string excelPath = "LR5-var9.xls";

        static void Main(string[] args)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            bool exit = false;



            while (!exit)
            {
                Console.WriteLine("\n=== Меню приложения (Гостиница) ===");
                Console.WriteLine("1. Загрузить базу из Excel (файл: " + excelPath + ")");
                Console.WriteLine("2. Просмотр таблиц");
                Console.WriteLine("3. Удалить элемент по ключу");
                Console.WriteLine("4. Добавить элемент");
                Console.WriteLine("5. Выполнить запросы (4 запроса)");
                Console.WriteLine("6. Сохранить изменения в Excel");
                Console.WriteLine("0. Выход");
                Console.Write("Выбор: ");
                string key = Console.ReadLine();

                try
                {
                    switch (key)
                    {
                        case "1":
                            Console.Write("Путь к файлу (Enter - использовать стандартный): ");
                            var inputPath = Console.ReadLine();
                            if (!string.IsNullOrWhiteSpace(inputPath))
                                excelPath = inputPath.Trim();
                            else excelPath = "LR5-var9.xls";
                                HotelService.LoadFromExcel(excelPath);
                            Console.WriteLine("Загружено: Clients={0}, Rooms={1}, Bookings={2}",
                                HotelService.Clients.Count, HotelService.Rooms.Count, HotelService.Bookings.Count);
                            break;

                        case "2":
                            ShowAllTables();
                            break;

                        case "3":
                            DeleteItemMenu();
                            break;

                        case "4":
                            AddItemMenu();
                            break;

                        case "5":
                            RunQueriesMenu();
                            break;

                        case "6":
                            HotelService.SaveToExcel(excelPath);
                            Console.WriteLine("Сохранено в " + excelPath);
                            break;

                        case "0":
                            exit = true;
                            break;

                        default:
                            Console.WriteLine("Неверный выбор.");
                            break;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Ошибка: " + ex.Message);
                }
            }
        }

        static void ShowAllTables()
        {
            Console.WriteLine("\n-- Клиенты --");
            foreach (var c in HotelService.ViewClients())
                Console.WriteLine(c);

            Console.WriteLine("\n-- Номера --");
            foreach (var r in HotelService.ViewRooms())
                Console.WriteLine(r);

            Console.WriteLine("\n-- Бронирования --");
            foreach (var b in HotelService.ViewBookings())
                Console.WriteLine(b);
        }

        static void DeleteItemMenu()
        {
            Console.WriteLine("Удалить: 1-Клиент, 2-Номер, 3-Бронь");
            var sel = Console.ReadLine();
            Console.Write("Введите ID: ");
            if (!int.TryParse(Console.ReadLine(), out int id))
            {
                Console.WriteLine("Неверный ID.");
                return;
            }

            bool ok = false;
            switch (sel)
            {
                case "1":
                    ok = HotelService.DeleteClientById(id);
                    break;
                case "2":
                    ok = HotelService.DeleteRoomById(id);
                    break;
                case "3":
                    ok = HotelService.DeleteBookingById(id);
                    break;
                default:
                    Console.WriteLine("Неверный выбор.");
                    return;
            }

            Console.WriteLine(ok ? "Удалено." : "Элемент с таким ID не найден.");
        }

        static void AddItemMenu()
        {
            Console.WriteLine("Добавить: 1-Клиент, 2-Номер, 3-Бронь");
            var sel = Console.ReadLine();
            try
            {
                switch (sel)
                {
                    case "1":
                        Console.Write("ID: "); int cid = int.Parse(Console.ReadLine());
                        Console.Write("Фамилия: "); string last_name = Console.ReadLine();
                        Console.Write("Имя: "); string first_name = Console.ReadLine();
                        Console.Write("Отчество: "); string surname = Console.ReadLine();
                        Console.Write("Адрес: "); string addr = Console.ReadLine();
                        HotelService.AddClient(new Client(cid, last_name, first_name, surname, addr));
                        Console.WriteLine("Клиент добавлен.");
                        break;

                    case "2":
                        Console.Write("ID: "); int rid = int.Parse(Console.ReadLine());
                        Console.Write("Этаж: "); int floor = int.Parse(Console.ReadLine());
                        Console.Write("Число мест: "); int cap = int.Parse(Console.ReadLine());
                        Console.Write("Цена за сутки: "); decimal price = decimal.Parse(Console.ReadLine(), CultureInfo.InvariantCulture);
                        Console.Write("Категория: "); int cat = int.Parse(Console.ReadLine());
                        HotelService.AddRoom(new Room(rid, floor, cap, price, cat));
                        Console.WriteLine("Номер добавлен.");
                        break;

                    case "3":
                        Console.Write("ID брони: "); int bid = int.Parse(Console.ReadLine());
                        Console.Write("ID клиента: "); int bcid = int.Parse(Console.ReadLine());
                        Console.Write("ID номера: "); int brid = int.Parse(Console.ReadLine());
                        Console.Write("Дата брони (yyyy-MM-dd): "); DateTime bdate = DateTime.Parse(Console.ReadLine());
                        Console.Write("Заезд (yyyy-MM-dd): "); DateTime ci = DateTime.Parse(Console.ReadLine());
                        Console.Write("Выезд (yyyy-MM-dd): "); DateTime co = DateTime.Parse(Console.ReadLine());
                        HotelService.AddBooking(new Booking(bid, bcid, brid, bdate, ci, co));
                        Console.WriteLine("Бронь добавлена.");
                        break;

                    default:
                        Console.WriteLine("Неверный выбор.");
                        break;
                }
            }
            catch (FormatException)
            {
                Console.WriteLine("Ошибка формата ввода.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Невозможно добавить: " + ex.Message);
            }
        }

        static void RunQueriesMenu()
        {
            Console.WriteLine("\n--- Запросы ---");

            // 1 таблица
            Console.Write("A) Введите категорию для вывода всех номеров с этой категорией: ");
            int cat = int.Parse(Console.ReadLine());
            var rooms = HotelService.GetRoomsByCategory(cat);
            Console.WriteLine("\nНайденные номера:");
            foreach (var r in rooms)
                Console.WriteLine(r);

            // 2 таблицы
            Console.Write("\nB) Введите ID клиента для подсчёта бронирований: ");
            if (int.TryParse(Console.ReadLine(), out int clientId))
            {
                int cnt = HotelService.GetBookingsCountByClientId(clientId);
                Console.WriteLine($"Количество бронирований клиента {clientId}: {cnt}");
            }
            else
            {
                Console.WriteLine("Неверный ID.");
            }

            // 3 таблицы
            Console.WriteLine("\nC) Перечень всех бронирований с клиентом и номером (Query C):");
            var det = HotelService.GetBookingsDetailed();
            foreach (var x in det)
                Console.WriteLine(x);

            // 3 таблицы
            decimal revenue = HotelService.GetTotalRevenue();
            Console.WriteLine($"\nD) Общая предполагаемая выручка (Query D): {revenue}");
        }


    }
}



