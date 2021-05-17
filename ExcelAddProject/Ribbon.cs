using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration.CustomUI;

namespace ExcelAddProject
{
    [ComVisible(true)]
    public class RibbonController : ExcelRibbon
    {
        public override string GetCustomUI(string RibbonID)
        {
            return @"
      <customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
      <ribbon>
        <tabs>
          <tab id='tab1' label='Sharp Welder'>
            <group id='Test' label='Test Functions'>
              <button id='button1' imageMso='Coffee' label='First Button' onAction='OnButtonPressed' size='large'/>
              <button id='Weld_Data' imageMso='ModuleInsert' label='Прересчет WB' onAction='OnButtonPressed2' size='large'/>
              <button id='Print_Rates' imageMso='Chart3DPieChart' label='Посчитать брак' onAction='OnButtonPressed3' size='large'/>
            </group >
          </tab>
        </tabs>
      </ribbon>
    </customUI>";
        }
        
        public void OnButtonPressed(IRibbonControl control)
        {
            //Form1();
        }
        public void OnButtonPressed2(IRibbonControl control)
        {
            RepairRates.WeldData();
        }
        public void OnButtonPressed3(IRibbonControl control)
        {
            RepairRates.PrintRates();
        }
    }
}
