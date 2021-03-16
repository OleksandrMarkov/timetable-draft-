using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;



namespace TimetableApp
{
    class ExcelFileCheck
    {
        string file;



        public ExcelFileCheck(ExcelFile file)
        {
            this.file = file.FileName;
        }

        public string File
         {
                get
                    {
                        return file;
                    }
                set
                    {
                        file = value;
                    }
         }
    }
}
