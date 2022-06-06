using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddProject
{
    public class RatesContainerCOK
    {
        public double Overall = 0;
        public double NDEOverall = 0;
        public double NDEDone = 0;
        public double NDEReject = 0;
    }

    public class RatesContainer
    {
        public double Accept = 0;
        public double Repair = 0;
        public double Cutout = 0;
        public double Overall = 0;
        public double NDETotal()
        {
            double NDETotal = Repair + Cutout+ Accept;
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
        public List<KSS> QualKSS = new List<KSS>();
        public List<KSS> TestKSS = new List<KSS>();
        public List<RatesContainer> WelderRates = new List<RatesContainer>(16);
        public List<TimelineRates> TimelineRates = new List<TimelineRates>();
        public List<QualSimp> QualSimpl = new List<QualSimp>();
        public Welder()
        {
            for (int i = 0; i < 16; i++)
            {
                WelderRates.Add(new RatesContainer());
            }
        }
    }
    public class TimelineRates
    {
        public RatesContainer RatesContainer = new RatesContainer();
        public bool isQual = false, isGapQual = false, isRequal = false;
        public DateTime StartDate = new DateTime();
        public DateTime EndDate = new DateTime();

    }
    public class DateRange
    {
        public DateTime StartDate = new DateTime();
        public DateTime EndDate = new DateTime();    
    }
         
    public class Weld
    {
        public object ISONum, DrawingNum, WeldNumber, WeldMaterial, RTProtNum, UTProtNum, WeldProcess, Result, RTDate, UTDate, Object;
        public bool NDEcontrol, IsRepair;
        public List<string> Welders, WeldersToBlame = new List<string>();
        public double WeldDiam, WeldThick, WeldDiamInch;
        public DateTime EndDate = new DateTime();
    }
    public class KSS
    {
        public string WelderName;
        public string Stamp;
        public string KSSNumber;
        public double Diameter;
        public Thickness Thickness = new Thickness();
        public string Material;
        public string MaterialHeat;
        public string MaterialGroup;
        public string Position;
        public string KSSProcess;
        public string WelderProcess;
        public string WeldLayers;
        public string ObjectType;
        public string WeldType;
        public string WPS;
        public DateTime StartDate = new DateTime();
        public DateTime EndDate = new DateTime();
        public string VTrequestNumber;
        public DateTime VTrequestDate = new DateTime();
        public string VTprotocolNumber;
        public DateTime VTprotocolDate = new DateTime();
        public string VTresult;
        public string NDTrequestNumber;
        public DateTime NDTrequestDate = new DateTime();
        public string NDTprotocolNumber;
        public DateTime NDTprotocolDate = new DateTime();
        public string NDTType;
        public string NDTresult;
        public string MECHrequestNumber;
        public DateTime MECHrequestDate = new DateTime();
        public string MECHprotocolNumber;
        public DateTime MECHprotocolDate = new DateTime();
        public string StringBendresult;
        public string Hardnessresult;
        public string Impactresult;
        public string Dlstatus;
        public string DLname;
        public DateTime Dldate = new DateTime();
        public double Diametermm;
        public string Diameterinch;
        public string Material1;
        public string MaterialHeat1;
        public object MaterialPNo1;
        public string MaterialGroup1;
        public string Material2;
        public string MaterialHeat2;
        public string MaterialPNo2;
        public string MaterialGroup2;
        public string Prosedure;
        public string Tensile;
        public string TensileWM;
        public string Bend;
        public string Macro;
    }
    public class Thickness
    {
        public string EnterThickness;
        public double WelderThick()
        {
            char open = '(';
            char close = ')';
            if (EnterThickness.Contains(open) & EnterThickness.Contains(close))
            {
                return Convert.ToDouble(EnterThickness.Substring(EnterThickness.IndexOf(open) + 1, EnterThickness.IndexOf(open)-EnterThickness.IndexOf(close)-1)) ;
            }
            else
            return Convert.ToDouble(EnterThickness);
        }
        public double FullThick()
        {
            char open = '(';
            char close = ')';
            if (EnterThickness.Contains(open) & EnterThickness.Contains(close))
            {
                return Convert.ToDouble(EnterThickness.Substring(0, EnterThickness.Length - EnterThickness.IndexOf(open) - 1));
            }
            else
                return Convert.ToDouble(EnterThickness);
        }

    }
    public class QualSimp
    {
        public string WelderName;
        public string Stamp;
        public string DLname;
        public DateTime Dlstart;
        public DateTime Dlend;
        public string KSSProcess;
        public string WelderProcess;
        public string WeldLayers;
        public string Position;
        public QualDimentions QualDimentions = new QualDimentions();
        public QualPositions QualPositions = new QualPositions();
        public QualMaterials QualMaterials = new QualMaterials();
        public string MaterialGroup;
        public string ObjectType;
    }
    public class Qualification
    {
        public string WelderName;
        public string Stamp;
        public string DLname;
        public DateTime Dldate;
        public string KSSProcess;
        public string WelderProcess;
        public string WeldLayers;
        public KSS KSS = new KSS();
        public QualDimentions QualDimentions = new QualDimentions();
        public QualPositions QualPositions = new QualPositions();
        public QualMaterials QualMaterials = new QualMaterials();
        //public string VTprotocolNumber;
        //public DateTime VTprotocolDate;
        //public string NDTprotocolNumber;
        //public DateTime NDTprotocolDate;
        //public string NDTType;
        //public string MECHprotocolNumber;
        //public DateTime MECHprotocolDate;
        //public string StringBendresult;
        //public string Hardnessresult;
        //public string Impactresult;
        //public string TPB309number;
        //public DateTime TPB309date;
        public string ObjectType;
    }
    public class QualDimentions
    {
        public double DiameterMin;
        public double DiameterMax;
        public double ThiknessMin;
        public double ThiknessMax;
        //public QualDimentions GetQualDim(double Diameter, double Thikness, string Procedure)
        //    {        
        //    }
    }
    public class QualPositions
    {
        public string QualPosition;
        public string QualPositionsString;
        public string[] QualifacatedPositions(string KSSType, string WeldType, string QualType)
        {
            string[] QualArray = new string[5]; 
            switch (QualType)
            {
                case "ASME":
                    switch (KSSType)
                    {
                        case "Pipe":
                            switch (WeldType)
                            {
                                case "Groove":
                                    switch (QualPosition)
                                    {
                                        case "1G":
                                            QualArray = new string[] { "F" };
                                            break;
                                        case "2G":
                                            QualArray = new string[] { "F", "H" };
                                            break;
                                        case "5G":
                                            QualArray = new string[] { "F", "V", "O" };
                                            break;
                                        case "6G":
                                            QualArray = new string[] { "F", "V", "H", "O" };
                                            break;
                                    }
                                    break;
                                case "Fillet":
                                    switch (QualPosition)
                                    {
                                        case "1F":
                                            QualArray = new string[] { "F" };
                                            break;
                                        case "2F":
                                            QualArray = new string[] { "F", "H" };
                                            break;
                                        case "2FR":
                                            QualArray = new string[] { "F", "H" };
                                            break;
                                        case "4F":
                                            QualArray = new string[] { "F", "H", "O" };
                                            break;
                                        case "5F":
                                            QualArray = new string[] { "F", "V", "H", "O" };
                                            break;
                                    }
                                    break;
                            }
                            break;
                        case "Plate":
                            switch (WeldType)
                            {
                                case "Groove":
                                    switch (QualPosition)
                                    {
                                        case "1G":
                                            QualArray = new string[] { "F" };
                                            break;
                                        case "2G":
                                            QualArray = new string[] { "F", "H" };
                                            break;
                                        case "3G":
                                            QualArray = new string[] { "F", "V" };
                                            break;
                                        case "4G":
                                            QualArray = new string[] { "F", "O" };
                                            break;
                                    }
                                    break;
                                case "Fillet":
                                    switch (QualPosition)
                                    {
                                        case "1F":
                                            QualArray = new string[] { "F" };
                                            break;
                                        case "2F":
                                            QualArray = new string[] { "F", "H" };
                                            break;
                                        case "3F":
                                            QualArray = new string[] { "F", "H", "V" };
                                            break;
                                        case "4F":
                                            QualArray = new string[] { "F", "H", "O" };
                                            break;
                                    }
                                    break;
                            }
                            break;
                    }
                    break;
                case "AWS":
                    switch (KSSType)
                    {
                        case "Pipe":
                            switch (WeldType)
                            {
                                case "Groove":
                                    switch (QualPosition)
                                    {
                                        case "1G":
                                            QualArray = new string[] { "F" };
                                            break;
                                        case "2G":
                                            QualArray = new string[] { "F", "H" };
                                            break;
                                        case "5G":
                                            QualArray = new string[] { "F", "V", "OH" };
                                            break;
                                        case "6G":
                                            QualArray = new string[] { "F", "V", "H", "OH" };
                                            break;
                                    }
                                    break;
                                case "Fillet":
                                    switch (QualPosition)
                                    {
                                        case "1F":
                                            QualArray = new string[] { "F" };
                                            break;
                                        case "2F":
                                            QualArray = new string[] { "F", "H" };
                                            break;
                                        case "4F":
                                            QualArray = new string[] { "F", "H", "OH" };
                                            break;
                                        case "5F":
                                            QualArray = new string[] { "F", "H", "OH" };
                                            break;
                                    }
                                    break;
                            }
                            break;
                        case "Plate":
                            switch (WeldType)
                            {
                                case "Groove":
                                    switch (QualPosition)
                                    {
                                        case "1G":
                                            QualArray = new string[] { "F" };
                                            break;
                                        case "2G":
                                            QualArray = new string[] { "F", "H" };
                                            break;
                                        case "3G":
                                            QualArray = new string[] { "F", "V", "H" };
                                            break;
                                        case "4G":
                                            QualArray = new string[] { "F", "OH" };
                                            break;
                                    }
                                    break;
                                case "Fillet":
                                    switch (QualPosition)
                                    {
                                        case "1F":
                                            QualArray = new string[] { "F" };
                                            break;
                                        case "2F":
                                            QualArray = new string[] { "F", "H" };
                                            break;
                                        case "3F":
                                            QualArray = new string[] { "F", "H", "V" };
                                            break;
                                        case "4F":
                                            QualArray = new string[] { "F", "H", "OH" };
                                            break;
                                    }
                                    break;                                
                            }
                            break;
                    }                    
                    break;
                default:
                    QualArray = new string[] { "error" };
                    break;
            }
            return QualArray;
        }
        
    }
    public class QualMaterials
    {
        public string QualMaterial;    
    }
    public class WPS
    {
        public string Type;
        public string TCOSpec;
        public string Standart;
        public string MaterialScope;
        public string Name;
        public string DCCName;
        public string Revision;
        public DateTime ApproveDate = new DateTime();
        public DateTime RevisionDate = new DateTime();
        public string Status;
        public string WeldingProcess;
        public string Preheat;
        public string Interpass;
        public string PWHT;
        public List<FillerMaterial> FillerMaterial = new List<FillerMaterial>();
        List<string> TypeOfElements = new List<string>();
        List<string> GroupMaterial = new List<string>();
        List<string> TypeOfJoint = new List<string>();
        public QualDimentions QualDimentions = new QualDimentions();
        public QualMaterials QualMaterials = new QualMaterials();
    }
    public class FillerMaterial
    {
        string AWSSpec;
        string AWSClass;
        string Name;
        string Size;    
    }
    public class Tools
    { 
    
    }
}
