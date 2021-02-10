using System;
using System.IO;

namespace LexTalionis.DocXTools
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var template = @"e:\!Projects\My Src\!Git\DocXTools\DocXTools.Integration\LeaveOfAbsenceOrder.docx";
            var result = @"e:\!Projects\My Src\!Git\DocXTools\DocXTools.Integration\LeaveOfAbsenceOrder_dev.docx";
            File.Copy(template, result, true);
            using (var word = Word.Open(result))
            {
                word.DeleteLastParagraph();
            }

            Console.WriteLine("Final");
            Console.ReadKey();
        }
    }
}
