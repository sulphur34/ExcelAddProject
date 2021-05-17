using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddProject
{
    public class RatesContainer
    {
        public int Accept = 0;
        public int Repair = 0;
        public int Cutout = 0;
        public int Overall = 0;
        public double NDETotal()
        {
            
            int NDETotal = Repair + Cutout+ Accept;
            return NDETotal;
        }
        public object Rates()
        {
            if (NDETotal() == 0)
            {
                return "-";
            }
            else
            {
                double Rates = (Repair + Cutout) / NDETotal();
                return Rates;
            }
        }
        public object NDErates()
        {
            if (Overall == 0)
            {
                return "-";
            }
            else
            {
                double NDErates = NDETotal() / Overall;
                return NDErates;
            }
            
        }       
    }
    public class Welder
    {
        public object WeldersName, Stamp;
        public List <Weld> WelderWelds = new List<Weld>();
        public List<RatesContainer> WelderRates = new List<RatesContainer>(13);
        public Welder()
        {
            for (int i = 0; i < 13;i++)
            {
                WelderRates.Add(new RatesContainer());
            }
            
        }
    }
    public class Weld
    {
        public object ISONum, DrawingNum, WeldNumber, WeldMaterial, RTProtNum, UTProtNum, WeldProcess, Result, RTDate, UTDate, Object;
        public List<string> Welders, WeldersToBlame = new List<string>();
        public double WeldDiam, WeldThick;
        public DateTime EndDate = new DateTime();
    }
}
