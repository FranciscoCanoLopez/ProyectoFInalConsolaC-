using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProyectoFInalConsolaC_
{
    internal class CsvRow
    {
        public Dictionary<string, string> Fields { get; set; } = new();
        public string? Player { get; set; }
        public string? Nationality { get; set; }
        public string? Position { get; set; }
        public string? Club { get; set; }
        public string? Age { get; set; }
        public string? Matches { get; set; }
        public string? Goals { get; set; }
    }
}
