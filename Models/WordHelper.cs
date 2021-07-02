using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace MyTest.Models
{
    public class WordHelper
    {
        #region Declaration

        IDictionary<string, string> Items { get; }
        IDictionary<string, string> ItemsTable { get; }

        readonly Application _app;

        static readonly Regex _regex;

        const string MARK = "__$__";

        #endregion

        #region Constructors

        public WordHelper()
        {
            _app = new Application();

            Items = new Dictionary<string, string>();
            ItemsTable = new Dictionary<string, string>();

            AppDomain.CurrentDomain.ProcessExit += (s, e) =>
            {
                foreach (Document doc in _app.Documents)
                {
                    doc.Close(SaveChanges: WdSaveOptions.wdDoNotSaveChanges);
                }
                _app.Quit();
            };
        }

        static WordHelper()
        {
            _regex = new Regex(@"\$[\{,\[](.+)[\},\]]");
        }

        #endregion

        #region Public methods

        public void Logic(string file)
        {
            try
            {
                object missing = Type.Missing;

                _app.Documents.Open(file);

                SearchItems(@"$\{*\}", Items);
                SearchItems(@"$\[*\]", ItemsTable);

                foreach (var item in Items)
                {
                    Find find = _app.Selection.Find;
                    find.Text = item.Key;
                    Console.Write($"{GetNormalStr(item.Key)}: ");
                    find.Replacement.Text = Console.ReadLine();

                    object wrap = WdFindWrap.wdFindContinue;
                    object replace = WdReplace.wdReplaceAll;

                    find.Execute(FindText: Type.Missing,
                        Forward: true,
                        Wrap: wrap,
                        MatchWildcards: false,
                        MatchSoundsLike: missing,
                        ReplaceWith: missing,
                        Replace: replace
                        );
                }

                FillTable();

                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Введите путь сохранения файла (Без названия файла) или пропустите, чтобы сохранить на рабочий стол.");
                Console.ResetColor();
                Console.Write("Путь: ");
                var pathToSave = VerifityPath(Console.ReadLine());

                Console.Write("Выберите формат сохранения. 1 - pdf, 2 - docx: ");
                var format = Console.ReadLine();

                if (format != "1")
                {
                    _app.ActiveDocument.SaveAs2(pathToSave + DateTime.Now.ToString("HH.mm.ss dd.MM.yy") + " myDock.docx");
                }
                else
                {
                    _app.ActiveDocument.ExportAsFixedFormat(pathToSave + DateTime.Now.ToString("HH.mm.ss dd.MM.yy") + " myDock.pdf", WdExportFormat.wdExportFormatPDF);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"{ex.Message}");
            }
            finally
            {
                _app.ActiveDocument.Close(SaveChanges: WdSaveOptions.wdDoNotSaveChanges);
            }
            Console.WriteLine("End");
        }

        #endregion

        #region Private methods

        void SearchItems(string findText, IDictionary<string, string> items)
        {
            _app.ActiveDocument.Select();

            var result = _app.Selection.Find.Execute(findText, MatchWildcards: true);

            while (result)
            {
                items[_app.Selection.Text] = string.Empty;
                result = _app.Selection.Find.Execute(findText, MatchWildcards: true);
            }
        }
        void FillTable()
        {
            if (ItemsTable?.Keys?.Count > 0)
            {
                Console.WriteLine("\nЗаполните поля таблицы:\n");

                try
                {
                    _app.ActiveDocument.Select();
                    if (!_app.Selection.Find.Execute(ItemsTable.Keys.First())) return;

                    Microsoft.Office.Interop.Word.Range range = _app.Selection.Range;

                    var t = range.Tables[1];
                    var texts = new List<string>();

                    foreach (Row row in t.Rows)
                    {
                        foreach (Cell cell in row.Cells)
                        {
                            var verifity = cell.Range.Text.Contains("$[");

                            texts.Add(verifity ? MARK : TrimCellText(cell.Range.Text));

                            cell.Range.Text = string.Empty;
                        }
                    }

                    int curr = 1;

                    while (true)
                    {
                        var rList = new List<string>();

                        var items = ItemsTable.Keys.GetEnumerator();

                        foreach (var text in texts)
                        {
                            string aText;

                            if (text == MARK)
                            {
                                items.MoveNext();
                                Console.Write($"{GetNormalStr(items.Current)}: ");
                                aText = Console.ReadLine();
                            }
                            else aText = text;

                            rList.Add(aText);
                        }

                        Row row = t.Rows[curr];

                        for (int i = 0; i < row.Cells.Count; i++)
                        {
                            row.Cells[i + 1].Range.Text = rList[i];
                        }

                        Console.Write("Введите 1 - Добавить ещё, 2 - закончить: ");
                        var choice = Console.ReadLine();
                        if (choice != "1") break;
                        t.Rows.Add(Type.Missing);
                        curr++;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error: {ex.Message}");
                }
            }
        }
        static string GetNormalStr(string source) => _regex.Match(source)?.Groups[1]?.Value;
        static string TrimCellText(string str)
        {
            if (string.IsNullOrEmpty(str)) return str;

            var end = str[^2..];

            if (end == "\r\a")
            {
                return str.Remove(str.Length - 2);
            }

            return str;
        }
        static string VerifityPath(string path)
        {
            var defaultPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\";

            if (!Directory.Exists(path)) return defaultPath;
            var p = path.Replace("/", @"\");
            return p[^1] != '\\' ? p + '\\' : p;
        }

        #endregion
    }
}
