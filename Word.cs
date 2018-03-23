using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Table = LexTalionis.DocXTools.Table;

namespace LexTalionis.DocXTools
{
    /// <summary>
    /// Класс для работы с шаблонами Word
    /// </summary>
    public class Word : IDisposable
    {
        private static WordprocessingDocument _doc;
        private readonly Dictionary<string, string> _bookmarks = new Dictionary<string, string>();
        readonly List<Table> _tables = new List<Table>();
       
        /// <summary>
        /// Открыть шаблон
        /// </summary>
        /// <param name="stream">шаблон (поток)</param>
        /// <returns>шаблон</returns>
        public static Word Open(Stream stream)
        {
            _doc = WordprocessingDocument.Open(stream, true);
            var word = new Word();
            return word;
        }

        /// <summary>
        /// Открыть шаблон
        /// </summary>
        /// <param name="filename">путь к файлу</param>
        /// <returns>шаблон</returns>
        public static Word Open(string filename)
        {
            _doc = WordprocessingDocument.Open(filename, true);
            var word = new Word();
            return word;
        }

        /// <summary>
        /// Заполнить закладку
        /// </summary>
        /// <param name="key">имя закладки</param>
        /// <param name="value">значение</param>
        public void SetStatic(string key, string value)
        {
            _bookmarks.Add(key, value);
        }

        /// <summary>
        /// Заполнить закладки
        /// </summary>
        /// <param name="bookmarks">словарь закладок</param>
        public void SetStatic(Dictionary<string, string> bookmarks)
        {
            foreach (var item in bookmarks)
            {
                _bookmarks.Add(item.Key, item.Value);
            }
        }

        /// <summary>
        /// Добавить строку закладок в таблицу
        /// </summary>
        /// <param name="row">строка закладок</param>
        /// <param name="id">идентификатор таблицы (для возможности добавлять строки в несколько таблиц)</param>
        public void AddTableRow(Dictionary<string, string> row, byte id)
        {
            var ltable = _tables.FirstOrDefault(x => x.Order == id);
            if (ltable != null)
                ltable.Rows.Add(row);
            else
                _tables.Add(new Table
                    {
                        Order = id, 
                        Rows = new List<Dictionary<string, string>> { row }
                    });
        }

        /// <summary>
        /// Добавить строки в таблицу
        /// </summary>
        /// <param name="table">коллекция строк закладок</param>
        /// <param name="id">идентификатор таблицы (для возможности добавлять строки в несколько таблиц)</param>
        public void AddTableRow(List<Dictionary<string, string>> table, byte id)
        {
            var ltable = _tables.FirstOrDefault(x => x.Order == id);
            if (ltable != null)
                    ltable.Rows.AddRange(table);
            else
            {
                _tables.Add(new Table
                {
                    Order = id,
                    Rows = table
                });
            }
        }

        public void Dispose()
        {
            if (_bookmarks.Count > 0)
                FillBookmarks(_doc.MainDocumentPart.RootElement, _bookmarks);

                foreach (var table in _tables)
                {
                    FillTables(table.Rows);
                }
            _doc.Dispose();
        }

        private static void FillTables(List<Dictionary<string, string>> list)
        {
            var firstRow = list.FirstOrDefault();
            if (firstRow == null)
                return;
            var firstBookmark = firstRow.FirstOrDefault();
            var row = GetRow(firstBookmark.Key);
            if (row == null)
                return;
            var table = row.Parent;

            foreach (var bookmarks in list)
            {
                var newRow = (TableRow)row.Clone();
                FillBookmarks(newRow, bookmarks);
                table.AppendChild(newRow);
            }
            row.Remove();
        }

        private static TableRow GetRow(string key)
        {
            var tableKey =
               _doc.MainDocumentPart.RootElement.Descendants<BookmarkStart>().FirstOrDefault(x => x.Name == key);
            if (tableKey == null)
                return null;
            var paragraph = tableKey.Parent;
            var cell = paragraph.Parent;
            return (TableRow)cell.Parent;
        }

        private static void FillBookmarks(OpenXmlElement root, Dictionary<string, string> bookmarks)
        {
            foreach (var itemBookmark in bookmarks)
            {
                var start = root.Descendants<BookmarkStart>().FirstOrDefault(x => x.Name == itemBookmark.Key);
                if (start == null)
                    continue;
                var elem = start.NextSibling();
                var run = (Run)elem;
                while (elem != null && !(elem is BookmarkEnd))
                {
                    var nextElem = elem.NextSibling();
                    elem.Remove();
                    elem = nextElem;
                }

                run.GetFirstChild<Text>().Text = itemBookmark.Value;
                start.Parent.InsertAfter(run, start);
            }
        }

        /// <summary>
        /// Удалить строку таблицы
        /// </summary>
        /// <param name="key">закладка на строке таблицы</param>
        public void DelTableRow(string key)
        {
            var row = GetRow(key);
            if (row == null)
                return;
            row.Remove();
        }
        /// <summary>
        /// Удалить таблицу
        /// </summary>
        /// <param name="key">закладка в таблице</param>
        public void DelTable(string key)
        {
            var row = GetRow(key);
            var table = row.Parent;
            table.Remove();
        }
    }
}
