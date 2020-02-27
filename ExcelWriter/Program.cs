using System;

namespace ExcelWriter
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelWriter ex = new ExcelWriter();
            ex.readXLS(@"./ExcelTemplate/TemplateTest.xlsx");
            
        }
    }
}
