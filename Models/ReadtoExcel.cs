using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CopyFromallExcelFiles.Models
{
    public class ReadtoExcel
    {
        String? kundenavn = "";
        String? Varebeskrivelse = "";
        String? antalenheder = "";
        String? nettoenheder = "";
        String? prisialt = "";
        String? prisperenhed = "";
        List<VarerBeskrivelse> varerBeskrivelses = new List<VarerBeskrivelse>();
        List<VarerBeskrivelse> opryddetvarer = new List<VarerBeskrivelse>();
        List<VarerBeskrivelse> returnlist = new List<VarerBeskrivelse>();
        string[] navn = new string[41];
        string[] forkortelse = new string[41];
        public List<VarerBeskrivelse> ReadfromExcel()
        {         
            Console.WriteLine("Begin to read Excel file"); 
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var files = Directory.GetFiles(@"C:\Users\KOM\source\repos\CopyFromallExcelFiles\ExcelFiles\AC");
            foreach (var item in files)
            {
                byte[] bin = File.ReadAllBytes(item);
                using (MemoryStream stream = new MemoryStream(bin))
                {
                    using (ExcelPackage excelPackage = new ExcelPackage(stream))
                    {
                        foreach (ExcelWorksheet worksheet in excelPackage.Workbook.Worksheets)
                        {
                            for (int i = 2; i <= worksheet.Dimension.End.Row; i++)
                            {
                                kundenavn = worksheet.Cells[i, 2]?.Value?.ToString();
                                Varebeskrivelse = worksheet.Cells[i, 4]?.Value?.ToString();
                                antalenheder = worksheet.Cells[i, 5]?.Value?.ToString();
                                nettoenheder = worksheet.Cells[i, 6]?.Value?.ToString();
                                prisialt = worksheet.Cells[i, 7]?.Value?.ToString();
                                prisperenhed = worksheet.Cells[i, 8]?.Value?.ToString();
                 
                                VarerBeskrivelse test = new VarerBeskrivelse( "2021", "K2", "AC" ,kundenavn, Varebeskrivelse, null, null ,null, antalenheder, nettoenheder, prisialt, prisperenhed, null);
                                varerBeskrivelses.Add(test);

                            }
                        }
                    }
                }
            }
            

            foreach (var item in varerBeskrivelses)
            {
                if (item.Kundenavn != null)
                {
                        opryddetvarer.Add(item);
                }
                  
            }

            foreach (var item in opryddetvarer)
            {
                if (item.VareBeskrivelse.Contains("ØKO"))
                {
                    item.TypeProduction = "øko";
                }
                else
                {
                    item.TypeProduction = "konv";
                    
                }
                var output = item.VareBeskrivelse.Split('(').Last();
                var update = item.VareBeskrivelse.Split('(').First();
                var result = output.Split(')').First();
                item.Land = result.ToUpper(); 
                item.VareBeskrivelse = update;

                if (item.Prisperenhed.Equals("#DIV/0!"))
                {
                    item.Prisperenhed = "0";
                }

                if (item.Land == "0" || item.Land.Length < 2 || item.Land.Equals("SVAMPE BLANDEDE SHITAKE/MARK /ØSTH/PORTO"))
                {
                    item.Land = null;
                }

                if (item.VareBeskrivelse.Contains("Diverse"))
                {
                    item.VareBeskrivelse = null;
                }
            }

            ReadHørkarm();
            Updatename();
            returnlist = opryddetvarer.OrderBy(x => x.Kundenavn).ToList();

            return returnlist;
        }

        public void Updatename()
        {
            Console.WriteLine("Updatename");
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;  
            var bin = File.ReadAllBytes("C:\\Users\\KOM\\source\\repos\\CopyFromallExcelFiles\\ExcelFiles\\BC.xlsx"); 
            List<string> Navneliste= new List<string>();
            List<string> forkortelseliste = new List<string>();

            using (MemoryStream stream = new MemoryStream(bin))
            {
                using (ExcelPackage excelPackage = new ExcelPackage(stream))
                {  
                   ExcelWorksheet excelWorksheet2 = excelPackage.Workbook.Worksheets[1];

                    for (int i = 2; i < excelWorksheet2.Dimension.End.Row; i++)
                    {
                        var Fuldekundenavn = excelWorksheet2.Cells[i, 2]?.Value?.ToString(); 
                        var Forkortelse = excelWorksheet2.Cells[i, 3]?.Value?.ToString();

                        Navneliste.Add(Fuldekundenavn);
                        forkortelseliste.Add(Forkortelse);
                    }
                }
            }

            foreach (var forkortelse in forkortelseliste)
            {

                Console.WriteLine(forkortelse);
                foreach (var item in Navneliste)
                {
                    Console.WriteLine(item);
                    foreach (var cleanup in opryddetvarer)
                    {
                        for (int i = 0; i < opryddetvarer.Count; i++)
                        {
                            if (item.Contains(cleanup.Kundenavn[i]))
                            {
                                cleanup.Kundenavn = forkortelse;
                            }
                        }

                       
                    }
                }
            }

        }

        public void ReadHørkarm()
        {
            Console.WriteLine("Begin to read Hørkarm");
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var bin = File.ReadAllBytes("C:\\Users\\KOM\\source\\repos\\CopyFromallExcelFiles\\ExcelFiles\\Hørkram.xlsx");
            
            using (MemoryStream stream = new MemoryStream(bin))
            {
                using (ExcelPackage excelPackage = new ExcelPackage(stream))
                {
                    ExcelWorksheet excelWorksheet2 = excelPackage.Workbook.Worksheets[1];

                        for (int i = 2; i < excelWorksheet2.Dimension.End.Row; i++)
                        {
                            var kundenavn = excelWorksheet2.Cells[i, 2]?.Value?.ToString();
                            var varer = excelWorksheet2.Cells[i, 4]?.Value?.ToString();
                            var Salgshovedvaregruppe = excelWorksheet2.Cells[i, 5]?.Value?.ToString();
                            var konventional = excelWorksheet2.Cells[i, 7]?.Value?.ToString();
                            var land = excelWorksheet2.Cells[i, 8]?.Value?.ToString();
                            var numbers = excelWorksheet2.Cells[i, 10]?.Value?.ToString();
                            var price = excelWorksheet2.Cells[i, 11]?.Value?.ToString();
                            var totalweight = excelWorksheet2.Cells[i, 13]?.Value?.ToString();
                            var weightperunit = excelWorksheet2.Cells[i, 14]?.Value?.ToString();

                        VarerBeskrivelse varerBeskrivelse = new VarerBeskrivelse("2021", "K2", "Hørkram", kundenavn, varer, konventional , Salgshovedvaregruppe, land, numbers, totalweight, price, null ,weightperunit); 
                            opryddetvarer.Add(varerBeskrivelse);
                        }

                }
            }

            foreach (var item in opryddetvarer)
            {
                if (item.TypeProduction.Equals("J"))
                {
                    item.TypeProduction = "øko";
                }

                if (item.TypeProduction.Equals("N"))
                {
                    item.TypeProduction = "konv";
                }

    
            }

        }

        public void Insert(List<VarerBeskrivelse> varer)
        {
            Console.WriteLine("Begin to read Excel file");
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var bin = File.ReadAllBytes("C:\\Users\\KOM\\source\\repos\\CopyFromallExcelFiles\\ExcelFiles\\2021 K2 renset.xlsx");
            using (MemoryStream stream = new MemoryStream(bin))
            {
                using (ExcelPackage excelPackage = new ExcelPackage(stream))
                {
                    foreach (ExcelWorksheet worksheet in excelPackage.Workbook.Worksheets)
                    {
                        for (int i = 2; i < varer.Count(); i++)
                        {
                                worksheet.Cells[i, 1].Value = varer[i].Year;
                                worksheet.Cells[i, 2].Value = varer[i].Kvartal;
                                worksheet.Cells[i, 3].Value = varer[i].Kundenavn;
                                worksheet.Cells[i, 5].Value = varer[i].Leverandør;
                                worksheet.Cells[i, 6].Value = varer[i].VareBeskrivelse;
                                worksheet.Cells[i, 7].Value = varer[i].TypeProduction;
                                worksheet.Cells[i, 9].Value =  varer[i].Prisperenhed;
                                worksheet.Cells[i, 10].Value = varer[i].Prisialt;
                                worksheet.Cells[i, 11].Value = varer[i].NettoKg;
                                worksheet.Cells[i, 12].Value = varer[i].KiloPris;
                                worksheet.Cells[i, 13].Value = varer[i].Land;                  
                        }   

                    }  
                    excelPackage.SaveAs(@"C:\Users\KOM\source\repos\CopyFromallExcelFiles\ExcelFiles\2021 K2 renset.xlsx");
                }
            }
            Console.WriteLine("Done");
        }

    }
}
