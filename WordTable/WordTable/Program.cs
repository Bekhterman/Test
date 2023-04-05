using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Task = System.Threading.Tasks.Task;


namespace WordTable
{
    internal class Program
    {
        private static async Task Main(string[] args)
        {
            try
            {
                var konApi = new KonturAPI();
                await konApi.RunAsync();
                Console.ReadKey();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }
    }

    public class KonturAPI
    {
        private const string URL = "https://api.kontur.ru/dc.contacts/v1/cus";
        private const string REGION_CODE = "18";

        public async Task RunAsync()
        {
            try
            {
                // Отправляем запрос к API и получаем ответ
                var httpClient = new HttpClient();
                var response = await httpClient.GetAsync(URL);
                var json = await response.Content.ReadAsStringAsync();
                if(json != null)
                {
                    Console.WriteLine("Формируется документ. Это займёт некоторое время.");
                    // Десериализуем json в объекты C#
                    var customers = JsonConvert.DeserializeObject<Rootobject>(json);
                    // Отбираем только те контролирующие органы, у которых в поле региона указано знаение "18"
                    var filteredList = customers.cus.Where(c => c.region.Contains(REGION_CODE)).ToList();
                    // Сортируем список контролирующих органов по типу и по коду
                    var sortedList = filteredList.OrderBy(c => c.type).ThenBy(c => c.code).ToList();

                    // Создаём документ Word
                    Word.Application wordApp = new Word.Application();
                    Word.Document wordDoc = wordApp.Documents.Add();

                    // Вставляем в начало документа дату и время выполнения, количество контролирующих органов.
                    Word.Range rng = wordDoc.Range();
                    var currentTime = DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss");
                    rng.Text = $"Дата выполнения: {currentTime}\nКоличество контролирующих органов:\nОбщее - {sortedList.Count}\n";

                    // Считаем количество контролирующих органов по типам
                    var typeCount = new Dictionary<string, int>();
                    foreach (var customer in sortedList)
                    {
                        if (!typeCount.ContainsKey(customer.type))
                        {
                            typeCount.Add(customer.type, 1);
                        }
                        else
                        {
                            typeCount[customer.type]++;
                        }
                    }

                    // Добавляем количество контролирующих органов по типам в документ Word
                    foreach (var type in typeCount)
                    {
                        Word.Range rng2 = wordDoc.Range();
                        rng2.End = wordDoc.Content.End; // Диапазон, в котором будет создана таблица (конец документа)
                        rng2.InsertAfter($"{type.Key} - {type.Value}\n");
                    }

                    // Добавим новую таблицу
                    Word.Range rng3 = wordDoc.Range(wordDoc.Content.End - 1);// Диапазон, в котором будет создана таблица (конец документа)
                    Word.Table table = wordDoc.Tables.Add(
                        rng3,
                        sortedList.Count + 1, // Количество строк
                        5 // Количество столбцов
                    );
                    table.Rows.Alignment = WdRowAlignment.wdAlignRowCenter; // Выраниваем таблицу по центру
                    table.Borders.Enable = 1; // Видимые рамки для таблицы

                    // Зададим ширину столбцов
                    table.Columns[1].Width = 50f; // Порядковый номер
                    table.Columns[2].Width = 50f; // Тип
                    table.Columns[3].Width = 70f; // Код
                    table.Columns[4].Width = 200f; // Имя
                    table.Columns[5].Width = 80f; // ИНН

                    // Заполним название столбцов
                    table.Cell(1, 1).Range.Text = "№";
                    table.Cell(1, 2).Range.Text = "Тип";
                    table.Cell(1, 3).Range.Text = "Код";
                    table.Cell(1, 4).Range.Text = "Имя";
                    table.Cell(1, 5).Range.Text = "ИНН";

                    int num = 2; // Заполняем таблицу со второй строки
                    foreach (var customer in sortedList)
                    {
                        table.Cell(num, 1).Range.Text = $"{num - 1}";
                        table.Cell(num, 2).Range.Text = $"{customer.type}";
                        table.Cell(num, 3).Range.Text = $"{customer.code}";
                        table.Cell(num, 4).Range.Text = $"{customer.name}";
                        table.Cell(num, 5).Range.Text = $"{customer.soun.inn}";
                        num++;
                    }

                    wordDoc.Save(); // Сохраняем документ в нужное место (для удобства)
                    wordDoc.Close();
                    wordApp.Quit();
                    Console.WriteLine("Документ успешно сформирован");
                }
                else
                    Console.WriteLine("Ошибка получения JSON");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }
    }

    //Класс для десериализации JSON
    public class Rootobject
    {
        public Cu[] cus { get; set; }
    }

    //Класс для десериализации JSON
    public class Cu
    {
        public string code { get; set; }
        public string name { get; set; }
        public string region { get; set; }
        public int timezone { get; set; }
        public string type { get; set; }
        public string[] keproviders { get; set; }
        public string flags { get; set; }
        public string inputwsversion { get; set; }
        public Soun soun { get; set; }
        public Certificates certificates { get; set; }
        public string comment { get; set; }
        public string[] uftkproviders { get; set; }
        public string[] newcodes { get; set; }
    }

    //Класс для десериализации JSON
    public class Soun
    {
        public string shortname { get; set; }
        public string fullname { get; set; }
        public DateTime validfrom { get; set; }
        public string inn { get; set; }
        public string kpp { get; set; }
        public string address { get; set; }
        public string document { get; set; }
        public string documentnum { get; set; }
        public DateTime documentdate { get; set; }
        public string phone { get; set; }
        public string email { get; set; }
        public DateTime validto { get; set; }
    }

    //Класс для десериализации JSON
    public class Certificates
    {
    }
}
