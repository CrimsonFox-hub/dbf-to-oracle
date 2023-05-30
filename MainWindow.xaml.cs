using Microsoft.Win32;
using System;
using System.Data;
using System.Data.Odbc;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Forms;

namespace dbf_to_oracle
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private OdbcConnection odbc; // ODBC коннект
        private FolderBrowserDialog FBD; //работа с проводником для выбора папки
        private Microsoft.Win32.OpenFileDialog openFile; // Работа с проводником для открытия файла
        private Microsoft.Win32.SaveFileDialog saveFile; // Работа с проводником для сохранения
        private OdbcCommand _command; // SQL комманда
        private OdbcDataReader _reader; // Для считывания
        private StreamWriter _fileStream; // Для работы с файлами и запись в них
        private DataTable dataTable; // Для получения схемы бд
        private StringBuilder str;  // Переменная для создания строк
        private readonly string _select = "SELECT * from "; // Запрос SQL
        private string dbfDirectory = string.Empty; // папка где хранится DBF файл
        private string dbfName = string.Empty; // Имя файла
        private string _namefile = string.Empty; // Имя без .DBF
        private string sqlDirectory = string.Empty; // путь до папки сохранения
        private string sqlName = string.Empty; // имя файла
        private string dbfPath = string.Empty; // путь к папке
        private string _columnName = string.Empty; // Имя колонки
        private string _columnType = string.Empty; // Тип колонки(int, varchar, DateTime)
        private string _insertText = string.Empty; // текст Insert операции

        /// <summary>
        /// действия при запуске
        /// </summary>
        public MainWindow()
        {
            InitializeComponent();
            _start.IsEnabled = false;
        }

        private void DBFLoad_Click(object sender, RoutedEventArgs e)
        {   
            FBD = new FolderBrowserDialog(); // создаем экземпляр FolderBrowserDialog

            if (FBD.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                dbfPath = FBD.SelectedPath;
                _DBF.Text = dbfPath; // записываем имя файла документа в TextBox 
                openFile = new Microsoft.Win32.OpenFileDialog(); // создаем экземпляр OpenFileDialog
                saveFile = new Microsoft.Win32.SaveFileDialog(); // создаем экземпляр SaveFileDialog
            }

            //перенос в старт с алгоритмом пробега по всем файлам
                if (openFile.ShowDialog() == true) // если окно проводника открыто и файл выбран
            {
                dbfDirectory = Path.GetDirectoryName(openFile.FileName); // записываем путь к файлу в переменую
                dbfName = openFile.SafeFileName; // записываем имя файла в переменную 
            }
            openFile.Reset();
            _namefile = dbfName.Remove(dbfName.IndexOf('.')); // получаем имя файла без расширения
            saveFile.FileName = _namefile + "_asd_4321.sql"; // Имя сохраняемого файла  = Имя DBF файла + расширение *.SQL
            if (saveFile.ShowDialog() == true) // окно сохранения 
            {
                sqlDirectory = dbfDirectory; // путь к файлу 
                sqlName = "\\" + saveFile.SafeFileName; // имя файла
            }
            saveFile.Reset(); 
            //

            if (dbfPath != null) 
            {
                _start.IsEnabled = true; // включаем кнопку конвертации 
            }
        }

        private void Start_Click(object sender, RoutedEventArgs e)
        {




            odbc = new OdbcConnection(@"Driver={Microsoft dBase Driver (*.dbf)}; SourceType=DBF;DefaultDir=" + dbfDirectory + ";Exclusive=No; Collate=Machine;NULL=NO;DELETED=NO; BACKGROUNDFETCH=NO"); // Подключение к DBF
            try
            {
                odbc.Open(); // Открываем подключение 
                _command = new OdbcCommand(_select + dbfName, odbc); // Отпавляем команду
                _reader = _command.ExecuteReader(); // Считываем ответ
                dataTable = _reader.GetSchemaTable(); // Считываем Схему таблицы 
                _fileStream = new StreamWriter(sqlDirectory + sqlName, true); // файл сохранения
                _fileStream.WriteLine("CREATE TABLE " + _namefile + "_asd_4321 ("); // Шаблон первой строки SQL
                str = new StringBuilder(); // Экземпляр StringBuilder 

                foreach (DataRow row in dataTable.Rows) // Прыгаем по всей схеме таблицы
                {
                    _columnName = row.Field<string>("ColumnName"); // Имя колонки в DBF файле
                    _columnType = new Operation().ConvertType(row.Field<Type>("DataType").Name); // Тип Колонки в DBF
                    str.Append(_columnName + "     " + _columnType + ", "); // добавляем все в StringBuilder 
                }

                var _newText = new Operation().AddLineString(str); // добавляем все в один текст через AddLineString разбиваем на строки
                _fileStream.WriteLine(_newText); // записываем в файл
                _fileStream.WriteLine(");"); // добавляем закрывающую скобку в последней строке 
                while (_reader.Read()) // Читаем
                {
                    for (int i = 0; i < _reader.FieldCount; i++) // Пока читаем
                    {
                        var ss = _reader[i]; // записываем во временную переменную данные
                        if (ss is double) // если тип double 
                        {
                            _insertText += ss + ", "; 
                        }
                        if (ss is string) // если тип string 
                        {
                            _insertText += "\'" + ss + "\', "; 
                        }
                        if (ss is DateTime) // если тип DateTime 
                        {
                            string date = ss.ToString(); 
                            DateTime _date = DateTime.Parse(date); 
                            date = _date.ToString("yyyy-MM-dd"); 

                            _insertText += "\'" + date + "\', "; 
                        }
                        if (ss.ToString() == "") //если пустые строки
                        {
                            _insertText += "\' \', ";
                        }
                        else 
                        {
                            _insertText += ss + ", "; 
                        }

                    }
                    _insertText = _insertText.TrimEnd(','); // удаляем запятую в конце
                    _fileStream.WriteLine("insert into " + _namefile + "_asd_4321 values (" + _insertText + ");"); // записываем в наш документ
                    _insertText = string.Empty; // обнуляемся
                }

                _fileStream.Close();
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(ex.Message);
            }
        }
    }
}
