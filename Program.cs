using Microsoft.Office.Interop.Excel;
using MySql.Data.MySqlClient;
using System;
using System.IO;
using System.Text;

namespace Excel2DB
{
    class Program
    {
        static void exceltocsv(string path, string pathcsv, char del)
        {
            //Создаем COM объект
            Application excelApp = new Application();

            //Проверяем существование Excel
            if (excelApp == null)
            {
                Console.WriteLine(DateTime.Now.ToLongTimeString() + "\t Excel не установлен");
                return;
            }

            //Определяем книгу, лист, размеры листа
            Workbook excelBook = excelApp.Workbooks.Open(@path);
            Worksheet excelSheet = (Worksheet)excelBook.Sheets[1];
            Range excelRange = excelSheet.UsedRange;

            //Определяем число строк и столбцов
            int rows = excelRange.Rows.Count;
            int cols = excelRange.Columns.Count;

            //Создание и удаление файла (чисто для теста)
            if (File.Exists(pathcsv))
            {
                File.Delete(pathcsv);
                string label = "Код" + del + "Наименование" + del + "Подразделение" + del + "Владелец процесса" + del + Environment.NewLine;
                File.WriteAllText(pathcsv, label, Encoding.GetEncoding(1251));
            }
            else
            {
                string label = "Код" + del + "Наименование" + del + "Подразделение" + del + "Владелец процесса" + del + Environment.NewLine;
                File.WriteAllText(pathcsv, label, Encoding.GetEncoding(1251));
            }

            //Чтение из файла
            for (int i = 5; i <= rows; i++)
            {
                //Через upstring идет сбор строки для *.csv
                string upstring = "";
                for (int j = 1; j <= cols; j++)
                {
                    //Проверяем непустые ячейки
                    if (excelRange.Cells[i, j] != null && excelRange.Cells[i, j].Value2 != null)
                    {
                        //Добавление в строку данных, в пустую(if) и непустую (else)
                        upstring = upstring + excelRange.Cells[i, j].Value2.ToString() + del;
                    }
                }

                //Проверяем подразделения и владельцев процессов - объединенные ячейки съезжают на следующую строку, присоединяем их обратно в нужную строку
                try
                {
                    excelRange.Cells[i + 1, 2].Value2.ToString();
                    upstring = upstring + del + Environment.NewLine;
                    File.AppendAllText(pathcsv, upstring, Encoding.GetEncoding(1251));
                    Console.WriteLine(DateTime.Now.ToLongTimeString() + "\t Строка добавлена в файл " + pathcsv);
                    upstring = "";
                }
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { }

            }
            //Выход из Excel
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            Console.WriteLine(DateTime.Now.ToLongTimeString() + "\t Запись в файл окончена");
        }

        static void db_structure()
        {
            //Отдельно прописываем название БД, т.к. изначально ее нет, позже будет создана
            string db = "test", table1 = "main", table2 = "process", table3 = "owners";

            //Объект для установления соединения с БД
            string conString = connectionString();
            MySqlConnection connection = new MySqlConnection(conString);

            //Открываем соединение
            connection.Open();

            //Запросы
            //Дропаем таблицу (для теста)
            string query = "DROP DATABASE IF EXISTS " + db;
            auto_query(query, connection);

            //Создали базу 'test', если не существует
            query = "CREATE DATABASE IF NOT EXISTS " + db;
            auto_query(query, connection);

            //Создание таблиц в 'test'
            //Таблица id - Название процесса
            query = "CREATE TABLE " + db + "." + table1 + "(id INT AUTO_INCREMENT, name TEXT, PRIMARY KEY (id))";
            auto_query(query, connection);

            //Таблица id = Код процесса
            query = "CREATE TABLE " + db + "." + table2 + "process (id INT AUTO_INCREMENT, process_code TEXT, PRIMARY KEY(id))";
            auto_query(query, connection);

            //Таблица id - подразделения и владельцы
            query = "CREATE TABLE " + db + "." + table3 + "owners (id INT AUTO_INCREMENT, retail BOOL, vip BOOL, oper BOOL, compliance BOOL, credit BOOL, corp BOOL, pro BOOL, mboard BOOL, sccenter BOOL, PRIMARY KEY(id))";
            auto_query(query, connection);
        }

        static void db_insert(string pathcsv, char del)
        {
            //Добавление строк
            using (StreamReader sr = new StreamReader(pathcsv, System.Text.Encoding.Default))
            {
                string line;
                sr.ReadLine();

                while ((line = sr.ReadLine()) != null)
                {
                    string process_code = "", process_name = "", query_insert_owner = "", query_insert_owner_values = "";
                    var str = line.Split(del);
                    for (int i = 0; i < str.Length; ++i)
                    {
                        //str[0] - Код процесса
                        //str[1] - Наименование процесса
                        //str[2] и далее - подразделения и владельцы процесса

                        //Отсеиваем подзаголовки таблицы
                        if (str[1] != "")
                            //Отсеиваем пустые ячейки
                            if (str[i] != "")
                                //Разбираем ячейки по переменным для вставки в запрос INSERT
                                switch (i)
                                {
                                    //Код процесса
                                    case 0:
                                        process_code = str[i];
                                        break;

                                    //Название процесса
                                    case 1:
                                        process_name = str[i];
                                        break;

                                    //Подразделения и владельцы процесса
                                    default:
                                        string owner = str[i];
                                        //По первым двум буквам определяем подразделение\владельца и добавляем в запрос INSERT
                                        switch (owner.Remove(2))
                                        {
                                            case "Re":
                                                query_insert_owner = query_insert_owner + "retail, ";
                                                query_insert_owner_values = query_insert_owner_values + "1, ";
                                                break;

                                            case "VI":
                                                query_insert_owner = query_insert_owner + "vip, ";
                                                query_insert_owner_values = query_insert_owner_values + "1, ";
                                                break;

                                            case "OP":
                                                query_insert_owner = query_insert_owner + "oper, ";
                                                query_insert_owner_values = query_insert_owner_values + "1, ";
                                                break;

                                            case "Co":
                                                query_insert_owner = query_insert_owner + "compliance, ";
                                                query_insert_owner_values = query_insert_owner_values + "1, ";
                                                break;

                                            case "Cr":
                                                query_insert_owner = query_insert_owner + "credit, ";
                                                query_insert_owner_values = query_insert_owner_values + "1, ";
                                                break;

                                            case "CO":
                                                query_insert_owner = query_insert_owner + "corp, ";
                                                query_insert_owner_values = query_insert_owner_values + "1, ";
                                                break;

                                            case "PR":
                                                query_insert_owner = query_insert_owner + "pro, ";
                                                query_insert_owner_values = query_insert_owner_values + "1, ";
                                                break;

                                            case "Ma":
                                                query_insert_owner = query_insert_owner + "mboard, ";
                                                query_insert_owner_values = query_insert_owner_values + "1, ";
                                                break;

                                            case "St":
                                                query_insert_owner = query_insert_owner + "scc, ";
                                                query_insert_owner_values = query_insert_owner_values + "1, ";
                                                break;
                                        }
                                        break;
                                }
                    }
                    try
                    {
                        //Формируем запрос
                        string query = "INSERT ";
                        Console.WriteLine(query_insert_owner + "\t" + query_insert_owner_values.Remove(query_insert_owner_values.Length - 2));
                    }
                    catch { }
                }
            }
        }
        static string connectionString()
        {
            string server = "localhost";
            string user = "user";
            string pword = "test";
            string SslMode = "none";
            string bd = "test";

            //Строка подключения к базе данных
            string connectionString = "server=" + server + ";user=" + user + ";password=" + pword + ";SslMode=" + SslMode;
            return connectionString;
        }

        static void auto_query(string query, MySqlConnection connection)
        {
            //Выполнение всех запросов
            MySqlCommand com = new MySqlCommand(query, connection);
            com.ExecuteNonQuery();
        }

        

        static void Main(string[] args)
        {
            Console.WriteLine(DateTime.Now.ToLongTimeString() + "\t Программа запущена");
            string temp = Environment.CurrentDirectory;
            string path = @"C:\Users\micha\Desktop\Энергосбыт\тестовые данные.xlsx";
            string pathcsv = "C:\\Users\\micha\\Desktop\\Энергосбыт\\тест.csv";
            char del = ';';

            //exceltocsv(path, pathcsv, del);
            //csvtodb(pathcsv, del);

            Console.WriteLine(DateTime.Now.ToLongTimeString() + "\t Работа программы окончена");
            Console.ReadLine();
        }
    }
}
