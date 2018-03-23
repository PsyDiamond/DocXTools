using System.Collections.Generic;

namespace LexTalionis.DocXTools
{
    class Table
    {
        public byte Order { get; set; }
        public List<Dictionary<string, string>> Rows = new List<Dictionary<string, string>>();
    }
}