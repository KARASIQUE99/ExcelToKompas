using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToKompas
{
    class CompRow<T> : IComparer<T>
       where T : Row
    {

        public int Compare(T x, T y)
        {
            int a = Convert.ToInt32(x.Count);
            int b = Convert.ToInt32(y.Count);
            if (a > b)
                return -1;
            if (b > a)
                return 1;
            else return 0;
        }
    }
    class Row
    {
        public string Unknown { get; set; } = "";
        public string Format { get; set; } = "";
        public string Zone { get; set; } = "";
        public string Position { get; set; } = "";
        public string Mark { get; set; } = "";
        public string Name { get; set; } = "";
        public string Count { get; set; } = "";
        public string Note { get; set; } = "";
        public string Mass { get; set; } = "";
        public string Material { get; set; } = "";
        public string User { get; set; } = "";
        public string Code { get; set; } = "";
        public string Factory { get; set; } = "";
        public string DocumentNumber { get; set; } = "";
        public string DocumentName { get; set; } = "";
        public string DocumentCode { get; set; } = "";
        public string CodeOKP { get; set; } = "";
        public string Type { get; set; } = "";

        public Row() { }

        public List<string> getRowAsList()
        {
            List<string> values = new List<string>();

            values.Add(Unknown);
            values.Add(Format);
            values.Add(Zone);
            values.Add(Position);
            values.Add(Mark);
            values.Add(Name);
            values.Add(Count);
            values.Add(Note);
            values.Add(Mass);
            values.Add(Material);
            values.Add(User);
            values.Add(Code);
            values.Add(Factory);
            values.Add(DocumentNumber);
            values.Add(DocumentName);
            values.Add(DocumentCode);
            values.Add(CodeOKP);

            return values;
        }
    }
}
