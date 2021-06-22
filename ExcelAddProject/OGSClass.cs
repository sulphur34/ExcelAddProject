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
        public int Overall = 0;
        public int NDEOverall = 0;
        public int NDEDone = 0;
        public int NDEReject = 0;
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
        public double WeldDiam, WeldThick;
        public DateTime EndDate = new DateTime();
    }
    public class KSS
    {
        public string WelderName;
        public string Stamp;
        public string KSSNumber;
        public double Diameter;
        public double Thikness;
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
        public string VTprotocolNumber;
        public DateTime VTprotocolDate;
        public string NDTprotocolNumber;
        public DateTime NDTprotocolDate;
        public string NDTType;
        public string MECHprotocolNumber;
        public DateTime MECHprotocolDate;
        public string StringBendresult;
        public string Hardnessresult;
        public string Impactresult;
        public string TPB309number;
        public DateTime TPB309date;
        public string ObjectType;
    }
    public class QualDimentions
    {
        public object DiameterMin;
        public object DiameterMax;
        public object ThiknessMin;
        public object ThiknessMax;
        //public QualDimentions GetQualDim(double Diameter, double Thikness, string Procedure)
        //    {        
        //    }
    }
    public class QualPositions
    {
        public string QualPosition;
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
}
