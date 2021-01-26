﻿using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.Remoting;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlPowerTools;
using TableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;

namespace LexTalionis.DocXTools
{
    /// <summary>
    /// Класс для работы с шаблонами Word
    /// </summary>
    public class Word : IDisposable
    {
        /// <summary>
        /// Документ
        /// </summary>
        private static WordprocessingDocument _doc;
        /// <summary>
        /// Закладки
        /// </summary>
        private readonly Dictionary<string, string> _bookmarks = new Dictionary<string, string>();
        /// <summary>
        /// Таблицы
        /// </summary>
        readonly List<Table> _tables = new List<Table>();

        private readonly List<string> _deleteFields = new List<string>(); 
       
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
            /* Защита от ошибок, если где то в строке используется null, то это вызывает проблемы */
            var val = value;
            if (val != null)
                val = val.Replace((char)0, ' ');
            _bookmarks.Add(key, val);
        }

        /// <summary>
        /// Заполнить закладки
        /// </summary>
        /// <param name="bookmarks">словарь закладок</param>
        public void SetStatic(Dictionary<string, string> bookmarks)
        {
            foreach (var item in bookmarks)
            {
                SetStatic(item.Key, item.Value);
            }
        }

        /// <summary>
        /// Удалить колекцию закладок
        /// </summary>
        /// <param name="list">закладки</param>
        public void DeleteStatic(IEnumerable<string> list)
        {
            _deleteFields.AddRange(list);
        }

        /// <summary>
        /// Удалить закладку
        /// </summary>
        /// <param name="item">имя закоадки</param>
        public void DeleteStatic(string item)
        {
            _deleteFields.Add(item);
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
        /// <summary>
        /// Завершить работу с закладками
        /// </summary>
        public void Dispose()
        {
            if (_bookmarks.Count > 0)
                FillBookmarks(_doc.MainDocumentPart.RootElement, _bookmarks);
            if (_deleteFields.Any())
                DeleteFields(_doc.MainDocumentPart.RootElement, _deleteFields);

                foreach (var table in _tables)
                {
                    FillTables(table.Rows);
                }
            _doc.Dispose();
        }

        private void DeleteFields(OpenXmlElement root, IEnumerable<string> deleteFields)
        {
            foreach (var item in deleteFields)
            {
                var bookmark = root.Descendants<BookmarkStart>().FirstOrDefault(x => x.Name == item);
                if (bookmark == null)
                    continue;
                var paragraph = bookmark.Parent;
                paragraph.Remove();
            }
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
            var sb = new StringBuilder();
            foreach (var itemBookmark in bookmarks)
            {
                var start = root.Descendants<BookmarkStart>().FirstOrDefault(x => x.Name == itemBookmark.Key);
                if (start == null)
                    continue;
                var elem = start.NextSibling();
                var run = elem as Run;
                while (elem != null && !(elem is BookmarkEnd))
                {
                    var nextElem = elem.NextSibling();
                    elem.Remove();
                    elem = nextElem;
                }
                if (run != null)
                    run.GetFirstChild<Text>().Text = itemBookmark.Value;
                else
                    sb.AppendLine(itemBookmark.Key);

                start.Parent.InsertAfter(run, start);
            }

            if (sb.Length > 0)
                throw new ServerException("Не верный дизайн документа, слеудет обратить внимание на закладки: " + sb);
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

        /// <summary>
        /// Объеденить отчеты в один
        /// </summary>
        /// <param name="reports">готовые отчеты для объединения</param>
        /// <returns>готовый отчет</returns>
        public static Stream MergeReports(IEnumerable<Stream> reports)
        {
            var source = new List<Source>();
            var i = 0;
            foreach (var item in reports)
            {
                ++i;
                using (item)
                {
                    var buffer = new byte[item.Length];
                    item.Read(buffer, 0, (int)item.Length);
                    source.Add(new Source(new WmlDocument(i.ToString(CultureInfo.InvariantCulture), buffer)));
                }
            }
            var tmp = Path.GetTempFileName();
            var mergedDoc = DocumentBuilder.BuildDocument(source);
            mergedDoc.SaveAs(tmp);

            return File.OpenRead(tmp);
        }
    }
}
