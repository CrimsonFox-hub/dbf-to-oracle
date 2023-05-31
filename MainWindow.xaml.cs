using Microsoft.Win32;
using System;
using System.Data;
using System.Data.Odbc;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Forms;
using DbfDataReader;


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
        private string dbfDirectory = string.Empty; // папка с DBF файлом
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
                _DBF.Text = dbfPath; // Заполняем имя файла документа в TextBox 
                // openFile и saveFile надо переделать так, чтобы они не открывали окно выбора, а брали/создавали файлы в фоне
                //openFile = new Microsoft.Win32.OpenFileDialog(); // создаем экземпляр OpenFileDialog
                //openFile.Filter = "dbf файлы (*.dbf)|*.dbf";
                //saveFile = new Microsoft.Win32.SaveFileDialog(); // создаем экземпляр SaveFileDialog
            }

            ////перенос отсюда
            //    if (openFile.ShowDialog() == true)
            //{
            //    dbfDirectory = dbfPath; // Добавляем путь к файлу в переменую
            //    dbfName = openFile.SafeFileName; // Добавляем имя файла в переменную 
            //}
            //openFile.Reset();
            //_namefile = dbfName.Remove(dbfName.IndexOf('.')); // получаем имя файла без расширения
            //saveFile.FileName = _namefile + "_asd_4321.sql"; // Имя сохраняемого файла  = Имя DBF файла + расширение *.SQL
            //if (saveFile.ShowDialog() == true)
            //{
            //    sqlDirectory = dbfPath; // путь к файлу 
            //    sqlName = "\\" + saveFile.SafeFileName; // имя файла
            //}
            //saveFile.Reset();
            //// и до сюда

            if (dbfPath != null) 
            {
                _start.IsEnabled = true;
            }
        }

        /// <summary>
        /// поиск всех dbf файлов
        /// </summary>
        /// <param name="Dir"></param>
        private void FileSearchFunction(object Dir)
        {
            DirectoryInfo DI = new DirectoryInfo((string)Dir);
            DirectoryInfo[] SubDir = null;

            try
            {
                SubDir = DI.GetDirectories();
            }
            catch
            {
                return;
            }

            for (int i = 0; i < SubDir.Length; ++i)
                this.FileSearchFunction(SubDir[i].FullName);

            FileInfo[] FI = DI.GetFiles();

            for (int i = 0; i < FI.Length; ++i)
                {
                    dbfName = FI[i].FullName;
                  //  _namefile = dbfName.Remove(dbfName.IndexOf('.')); // получаем имя файла без расширения
            };
            for (int i = 0; i < FI.Length; ++i)
            {
                _namefile = dbfName.Remove(dbfName.IndexOf('.')); // получаем имя файла без расширения
            };
            }



        private void Start_Click(object sender, RoutedEventArgs e)
        {
            var encoding = Encoding.GetEncoding(1251);

           // using var reader = new DbfDataReader.DbfDataReader(dbfPath);
            // odbc = new OdbcConnection(@"Driver={Microsoft dBase Driver (*.dbf)}; SourceType=DBF;DefaultDir=" + dbfPath + ";Exclusive=No; Collate=Machine;NULL=NO;DELETED=NO; BACKGROUNDFETCH=NO"); // Подключение к DBF
            try
            {
                odbc.Open(); // Открываем подключение 
                _command = new OdbcCommand(_select + dbfName, odbc); // Отпавляем команду
                _reader = _command.ExecuteReader(); // Считываем ответ
                dataTable = _reader.GetSchemaTable(); // Считываем Схему таблицы 
                _fileStream = new StreamWriter(dbfPath + _namefile + "_asd_4321", true); 
                _fileStream.WriteLine("CREATE TABLE " + _namefile + "_asd_4321 ("); // Шаблон первой строки SQL
                str = new StringBuilder(); // Экземпляр StringBuilder 

                if (dbfName != string.Empty && Directory.Exists(_DBF.Text))
                {

                    foreach (DataRow row in dataTable.Rows)
                    {
                        _columnName = row.Field<string>("ColumnName"); // Имя колонки в DBF файле
                        _columnType = new Operation().ConvertType(row.Field<Type>("DataType").Name); // Тип Колонки в DBF
                        str.Append(_columnName + "     " + _columnType + ", "); // добавляем все в StringBuilder 
                    }

                    var _newText = new Operation().AddLineString(str); // добавляем все в один текст через AddLineString разбиваем на строки
                    _fileStream.WriteLine(_newText);
                    _fileStream.WriteLine(");");
                    while (_reader.Read())
                    {
                        for (int i = 0; i < _reader.FieldCount; i++)
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
                        _insertText = _insertText.TrimEnd(',');
                        _fileStream.WriteLine("insert into " + _namefile + "_asd_4321 values (" + _insertText + ");");
                        _insertText = string.Empty;
                    }
                }
                _fileStream.Close();
                System.Windows.MessageBox.Show("Миграция завершена");
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(ex.Message);
            }
        }
    }
}
