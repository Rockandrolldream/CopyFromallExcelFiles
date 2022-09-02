using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CopyFromallExcelFiles.Models
{
    public class VarerBeskrivelse
    { 
        public string? Year { get; set; } 
        public string? Kvartal { get; set; } 
        public string? Leverandør { get; set; }
        public String? Kundenavn { get; set; } 
        public String? VareBeskrivelse { get; set; }     
        public String? TypeProduction { get; set; } 
        public String? VareKategori { get; set; } 
        public String? Land { get; set; }
        public String? Antalenheder { get; set; } 
        public String? NettoKg { get; set; } 
        public String? Prisialt { get; set; } 
        public String? Prisperenhed { get; set; } 

        public string? KiloPris { get; set; }

        public VarerBeskrivelse( string year ,string kvartal ,string leverandør ,string kundenavn, string vareBeskrivelse, string typeproduktion, string vareKategori ,string land, string antalenheder, string nettoKg, string prisialt, string prisperenhed, string kilopris)
        {
            Year = year;
            Kvartal = kvartal;
            Leverandør = leverandør;
            Kundenavn = kundenavn;
            VareBeskrivelse = vareBeskrivelse;
            TypeProduction = typeproduktion;
            VareKategori = vareKategori;
            Land = land;
            Antalenheder = antalenheder;
            NettoKg = nettoKg;
            Prisialt = prisialt;
            Prisperenhed = prisperenhed;
            KiloPris = kilopris;
        }
    }
}
