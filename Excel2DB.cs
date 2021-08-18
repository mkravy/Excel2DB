using Microsoft.Office.Interop.Excel;
using MySql.Data.MySqlClient;
using System;
using System.IO;
using System.Text;

namespace Excel2DB
{
    class Excel2DB
    {
        private static string query, db = "test", tree_table = "tree", table1 = "main", table2 = "process", table3 = "owners";
        private static string temp = Environment.CurrentDirectory, path = @"C:\Users\micha\Desktop\Энергосбыт\тестовые данные.xlsx", pathcsv = "C:\\Users\\micha\\Desktop\\Энергосбыт\\тест.csv";
        private static char del = ';';

        static void exceltocsv(string path, string pathcsv, char del)
        {
            //Создаем COM объект
            Application excelApp = new Application();

            //Проверяем существование Excel
            if (excelApp == null)
            {
                log("Excel не установлен");
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
            if (!File.Exists(pathcsv))
            {
                string label = "Код" + del + "Наименование" + del + "Подразделение" + del + "Владелец процесса" + del + Environment.NewLine;
                File.WriteAllText(pathcsv, label, Encoding.GetEncoding(1251));

                log("Создан файл " + pathcsv);
            }

            log("Начало чтения файла " + path);

            //Через upstring идет сбор строки для *.csv
            string upstring = "";

            try
            {
                //Чтение из файла
                for (int i = 5; i <= rows; i++)
                {
                    for (int j = 1; j <= cols; j++)
                    {
                        //Проверяем непустые ячейки
                        if (excelRange.Cells[i, j] != null && excelRange.Cells[i, j].Value2 != null)
                        {
                            //Вписываем подзаголовки на свои места
                            if (excelRange.Cells[i, 1].Value2 == null && excelRange.Cells[i, 2].Value2 != null)
                            {
                                upstring = upstring + del;
                            }
                            //Добавление в строку данных
                            upstring = upstring + excelRange.Cells[i, j].Value2.ToString() + del;
                        }
                    }

                    //Проверяем подразделения и владельцев процессов - объединенные ячейки съезжают на следующую строку, присоединяем их обратно в нужную строку
                    try
                    {
                        excelRange.Cells[i + 1, 2].Value2.ToString();
                        upstring = upstring + del + Environment.NewLine;
                        File.AppendAllText(pathcsv, upstring, Encoding.GetEncoding(1251));
                        upstring = "";

                        //log("Строка добавлена в файл " + pathcsv);
                    }
                    catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { }

                }
            }
            catch(Exception ex)
            {
                log_ex(ex, "Построчное чтение из файла");
            }
            
            //Выход из Excel
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

            log("Запись в файл " + pathcsv + " окончена");
        }

        static void db_structure(string db, string table1, string table2, string table3)
        {
            //Запросы
            //Создали базу 'test', если не существует
            try
            {
                query = "CREATE DATABASE IF NOT EXISTS " + db;
                auto_query(query);

                log("Создана база '" + db + "'");
            }
            catch (Exception ex)
            {
                log_query(ex, "\t База данных '" + db + "' не создана", query);
            }

            //Создание таблиц в 'test'
            //Таблица 1: id - Название процесса
            try
            {
                query = "CREATE TABLE " + db + "." + table1 + "(id INT AUTO_INCREMENT, name TEXT, PRIMARY KEY (id))";
                auto_query(query);

                log("Создана таблица " + table1);
            }
            catch (Exception ex)
            {
                log_query(ex, "\t Таблица " + table1 + " не создана", query);
            }

            //Таблица 2: id = Код процесса
            try
            {
                query = "CREATE TABLE " + db + "." + table2 + "(id INT AUTO_INCREMENT, process_code TEXT, PRIMARY KEY(id))";
                auto_query(query);

                log("Создана таблица " + table2);
            }
            catch (Exception ex)
            {
                log_query(ex, "\t Таблица " + table1 + " не создана", query);
            }

            //Таблица 3: id - подразделения и владельцы
            try
            {
                query = "CREATE TABLE " + db + "." + table3 + "(id INT AUTO_INCREMENT, retail BOOL, vip BOOL, oper BOOL, compliance BOOL, credit BOOL, corp BOOL, pro BOOL, mboard BOOL, sccenter BOOL, PRIMARY KEY(id))";
                auto_query(query);

                log("Создана таблица " + table3);
            }
            catch (Exception ex)
            {
                log_query(ex, "\t tаблица " + table1 + " не создана", query);
            }
        }

        static void db_insert(string pathcsv, char del, string db, string table1, string table2, string table3)
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
                                            query_insert_owner = query_insert_owner + "sccenter, ";
                                            query_insert_owner_values = query_insert_owner_values + "1, ";
                                            break;
                                    }
                                    break;
                            }
                    }

                    //Формируем запрос
                    //Добавление наименования процесса в таблицу 1
                    try
                    {
                        string query = "INSERT " + db + "." + table1 + "(name) VALUES ('" + process_name + "')";
                        auto_query(query);
                    }
                    catch (Exception ex)
                    {
                        log_query(ex, "\t Запрос INSERT в таблицу '" + table1 + "' не добавлен", query);
                    }

                    //Добавление кода процесса в таблицу 2
                    try
                    {
                        string query = "INSERT " + db + "." + table2 + "(process_code) VALUES ('" + process_code + "')";
                        auto_query(query);
                    }
                    catch (Exception ex)
                    {
                        log_query(ex, "\t Запрос INSERT в таблицу '" + table2 + "' не добавлен", query);
                    }

                    //Добавление владельцев и подразделения в таблицу 3
                    try
                    {
                        //Remove нужен, чтобы убрать последнюю запятую в запросе: (1,1,1,)
                        string query = "INSERT " + db + "." + table3 + "(" + query_insert_owner.Remove(query_insert_owner.Length - 2) + ") VALUES (" + query_insert_owner_values.Remove(query_insert_owner_values.Length - 2) + ")";
                        auto_query(query);
                    }
                    catch (System.ArgumentOutOfRangeException)
                    {
                        //Ошибка из-за того, что в эту таблицу нечего записать из соответствующей строки. Поэтому в первую колонку ставим 'NULL'
                        string query = "INSERT " + db + "." + table3 + "(retail) VALUES (NULL)";
                        auto_query(query);
                    }
                    catch (Exception ex)
                    {
                        log_query(ex, "\t Запрос INSERT в таблицу '" + table3 + "' не добавлен", query);
                    }
                }
            }
        }

        static string connectionString()
        {
            string server = "localhost";
            string user = "user";
            string pword = "test";
            string SslMode = "none";
            string db = "test";

            //Строка подключения к базе данных
            string connectionString = "server=" + server + ";user=" + user + ";password=" + pword + ";SslMode=" + SslMode;
            return connectionString;
        }

        static void auto_query(string query)
        {
            //Объект для установления соединения с БД
            string conString = connectionString();
            MySqlConnection connection = new MySqlConnection(conString);
            
            //Открываем соединение
            try
            {
                connection.Open();

                //log("Соединение с БД '" + db + "' установлено");
            }
            catch (Exception ex)
            {
                log_ex(ex, "\t Соединение с БД '" + db + "' не установлено");
            }

            //Выполнение запроса
            try
            {
                MySqlCommand com = new MySqlCommand(query, connection);
                com.ExecuteNonQuery();
                //log("Запрос " + query + " выполнен");
            }
            catch (Exception ex)
            {
                //log_ex(ex, "\t Запрос " + query + " не выполнен");
            }

            //Закрываем соединение
            try
            {
                connection.Close();

                //log("Соединение с БД '" + db + "' прекращено");
            }
            catch (Exception ex)
            {
                log_ex(ex, "\t Соединение с БД '" + db + "' не прекращено");
            }
        }

        static void log(string log_text)
        {
            string log_file = "C:\\Users\\micha\\Desktop\\Энергосбыт\\log.txt";
            string log_string = DateTime.Now.ToLongTimeString() + " " + log_text + "\n";
            Console.Write(log_string);
            File.AppendAllText(log_file, log_string, Encoding.GetEncoding(1251));
        }

        static void log_ex(Exception ex, string log_text)
        {
            //Выводит ошибку
            string log_file = "C:\\Users\\micha\\Desktop\\Энергосбыт\\log.txt";
            string log_string = "\n" + DateTime.Now.ToLongTimeString() + " " + log_text + "\n \t Ошибка: " + ex + "\n\n";
            Console.Write(log_string);
            File.AppendAllText(log_file, log_string, Encoding.GetEncoding(1251));
        }

        static void log_query(Exception ex, string log_text, string query)
        {
            //Кроме ошибки выводит сам запрос
            string log_file = "C:\\Users\\micha\\Desktop\\Энергосбыт\\log.txt";
            string log_string = "\n" + DateTime.Now.ToLongTimeString() + " " + log_text + "\n \t Ошибка: " + ex + "\n \t Запрос: " + query + "\n\n";
            Console.Write(log_string);
            File.AppendAllText(log_file, log_string, Encoding.GetEncoding(1251));
        }

        static void db_tree_structure(string db, string tree_table)
        {
            //Запросы
            try
            {
                //Создали базу 'test', если не существует
                query = "CREATE DATABASE IF NOT EXISTS " + db;
                auto_query(query);

                log("Создана база '" + db + "'");
            }
            catch (Exception ex)
            {
                log_query(ex, "\t База данных '" + db + "' уже создана", query);
            }

            //Создание таблицы в 'test'
            try
            {
                //Древоподобная таблица
                query = "CREATE TABLE " + db + "." + tree_table + "(id INT AUTO_INCREMENT, name TEXT, gen_num INT, path TEXT, parent_id INT, PRIMARY KEY (id))";
                auto_query(query);

                log("Создана таблица " + tree_table);
            }
            catch (Exception ex)
            {
                log_query(ex, "\t Таблица " + tree_table + " не создана", query);
            }
        }
        
        static void tree_structure(char del)
        {
            //Чтение строк
            using (StreamReader sr = new StreamReader(pathcsv, System.Text.Encoding.Default))
            {
                string line;
                sr.ReadLine();
                string name, gen_num, path, parent, parent_id="", word = "";

                while ((line = sr.ReadLine()) != null)
                {
                    var str = line.Split(del);
                    //str[0] - Код процесса
                    //str[1] - Наименование процесса
                    //str[2] и далее - подразделения и владельцы процесса

                    name = str[1];
                    path = str[0];

                    //Определяем букву для подзаголовков таблицы
                    try
                    {
                        word = path.Remove(1);
                    }
                    catch { }

                    //Добавляем букву подзаголовкам таблицы
                    if (path == "")
                        path = word;

                    //Определяем родителя элемента, чтобы в дальнейшем определить его id
                    try
                    {
                        parent = str[0].Remove(str[0].LastIndexOf('.'));
                        if (parent == "")
                            parent = str[0].Remove(str[0].LastIndexOf('.')+1);
                    }
                    catch
                    {
                        parent = "";
                    }

                    //Определяем номер элемента в его поколении
                    try
                    {
                        gen_num = str[0].Substring(str[0].Length - 1);
                    }
                    catch
                    {
                        gen_num = "";
                    }

                    if (gen_num == "")
                        gen_num = "0";

                    //Добавляем строки в таблицу 'tree'
                    insert_tree_structure(name, gen_num, path);

                    if (parent!="")
                    {
                        //Передаем значение 'parent_id'
                        parent_id = check_parent_id(parent);
                    }

                    //Добавляем значение 'parent_id'
                    insert_parent_id(path, parent_id);
                }
            }
        }

        static void insert_tree_structure(string name, string gen_num, string path)
        {
            //Формируем запрос
            try
            {
                //Добавление строки в таблицу
                string query = "INSERT " + db + "." + tree_table + $"(name, gen_num, path) VALUES('{name}', {gen_num}, '{path}')";
                auto_query(query);
            }
            catch (Exception ex)
            {
                log_query(ex, "\t Запрос INSERT в таблицу '" + tree_table + "", query);
            }
        }

        static void insert_parent_id(string path, string parent_id)
        {
            //Добавление 'parent_id' в таблицу
            try
            {
                string query = $"UPDATE {db}.{tree_table} SET parent_id = '{parent_id}' WHERE path = '{path}'";
                Console.WriteLine("query: " + query + "\n");
                auto_query(query);
            }
            catch (Exception ex)
            {
                log_query(ex, "\t Запрос INSERT в таблицу '" + tree_table + "", query);
            }
        }

        static string check_parent_id(string parent)
        {
            string conString = connectionString();
            MySqlConnection connection = new MySqlConnection(conString);
            string parent_id = "";

            //Открываем соединение
            try
            {
                connection.Open();

                //log("Соединение с БД '" + db + "' установлено");
            }
            catch (Exception ex)
            {
                log_ex(ex, "\t Соединение с БД '" + db + "' не установлено");
            }

            //Формируем запрос
            //Определяем id родителя
            try
            {
                string query = $"SELECT id FROM {db}.{tree_table} WHERE path = '{parent}'";
                Console.WriteLine(query);

                MySqlCommand com = new MySqlCommand(query, connection);
                parent_id = com.ExecuteScalar().ToString();
            }
            catch (Exception ex)
            {
                log_query(ex, "\t Запрос INSERT в таблицу '" + tree_table + "", query);
            }

            //Закрываем соединение
            try
            {
                connection.Close();

                //log("Соединение с БД '" + db + "' завершено");
            }
            catch (Exception ex)
            {
                log_ex(ex, "\t Соединение с БД '" + db + "' не завершено");
            }

            //Возвращаем parent_id
            return parent_id;
        }

        static void Main(string[] args)
        {
            string log_file = "C:\\Users\\micha\\Desktop\\Энергосбыт\\log.txt";
            if (File.Exists(log_file))
            {
                File.Delete(log_file);
            }
            else
            {
                log("Программа запущена");
            }

            //exceltocsv(path, pathcsv, del);
            //db_structure(db, table1, table2, table3);
            //db_insert(pathcsv, del, db, table1, table2, table3);
            //db_tree_structure(db, tree_table);
            tree_structure(del);

            log("Работа программы окончена");
            Console.ReadLine();
        }
    }
}
