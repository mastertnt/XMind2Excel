using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XMind2Xls
{
    public class ApplicationArguments
    {
        public string InputFile { get; set; }
        public string OutputFile { get; set; }

        public string Headers { get; set; }
        public bool Silent { get; set; }
        public bool RootAsWorksheet { get; set; }
    }
}
