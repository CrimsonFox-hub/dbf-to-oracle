using System;
using System.Text;

namespace dbf_to_oracle
{   
    /// <summary>
    /// класс для обработки значений
    /// </summary>
    internal class Operation
    {
        /// <summary>
        /// структура замены типов полей
        /// </summary>
        struct Type_
        {
            public const string int_ = "NUMBER";
            //добавить вариативность varchar(n)
            public const string varchar_ = "VARCHAR2(255)";
            public const string Datetime_ = "DATE";
        }

        private StringBuilder str_; // StringBuilder для работы с строками
        private string newText_ = string.Empty; // Для создания нового текста 

        /// <summary>
        /// Преобрузет типы DBF в типы SQL 
        /// </summary>
        /// <param name="type">DBF тип в формате string</param>
        /// <returns></returns>
        public string ConvertType(string type)
        {
            switch (type)
            {
                case "Double":
                    return Type_.int_; //Double to int
                case "String":
                    return Type_.varchar_; //String to varchar(255)
                case "DateTime":
                    return Type_.Datetime_; //DateTime, to datetime
            }
            return type;
        }

        /// <summary>
        /// заполнение данными таблиц при миграции
        /// </summary>
        /// <param name="paramter"></param>
        /// <returns></returns>
        public string AddLineString(StringBuilder paramter)
        {
            str_ = new StringBuilder();
            paramter.ToString().TrimEnd(',');
            string[] _text = paramter.ToString().Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < _text.Length; i++) { str_.AppendLine(_text[i] + ", "); }
            newText_ = str_.ToString();
            newText_ = newText_.TrimEnd('\n', '\r', ',');
            return newText_;
        }
    }
}
