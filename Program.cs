using CopyFromallExcelFiles.Models;
using OfficeOpenXml;

namespace CopyFromallExcelFiles
{
    internal class Program
    {
       
        static void Main(string[] args)
        {
            ReadtoExcel readtoExcel = new ReadtoExcel();
           
        List<VarerBeskrivelse> list = new List<VarerBeskrivelse>();
 
              list = readtoExcel.ReadfromExcel(); 
           

            readtoExcel.Insert(list);
        }
    }
}