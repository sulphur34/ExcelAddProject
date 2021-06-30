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
              <button id='button1' imageMso='DatabaseAccessBackEnd' label='Первые три товарных' onAction='OnButtonPressed' size='large'/>
              <button id='Weld_Data' imageMso='ModuleInsert' label='Прересчет WB' onAction='OnButtonPressed2' size='large'/>
              <button id='Print_Rates' imageMso='Chart3DPieChart' label='Посчитать брак' onAction='OnButtonPressed3' size='large'/>
              <button id='Print_Rates_COK' imageMso='DataSourceCatalogSOAP' label='Посчитать брак ЦОК' onAction='OnButtonPressed4' size='large'/>
              <button id='Print_Quals' imageMso='ReviewDeleteAllMarkupInPresentation' label='Квалификации' onAction='OnButtonPressed6' size='large'/>
              <button id='Print_Rates_Time' imageMso='ReviewDeleteAllMarkupInPresentation' label='Таймлайн' onAction='OnButtonPressed5' size='large'/>
            </group >
          </tab>
        </tabs>
      </ribbon>
    </customUI>";
        }
        
        public void OnButtonPressed(IRibbonControl control)
        {
            RepairRates.PrintFirstThree();
        }
        public void OnButtonPressed2(IRibbonControl control)
        {
            RepairRates.WeldData(true, true);
        }
        public void OnButtonPressed3(IRibbonControl control)
        {
            RepairForm FormRepair = new RepairForm();
            FormRepair.Show();
            //RepairRates.PrintRates();
        }
        public void OnButtonPressed4(IRibbonControl control)
        {
            COKReport COKReport = new COKReport();
            COKReport.Show();
            //RepairRates.PrintRates();
        }
        public void OnButtonPressed5(IRibbonControl control)
        {
            Timeline timeline = new Timeline();
            timeline.Show();
            //RepairRates.PrintRates();
        }
        public void OnButtonPressed6(IRibbonControl control)
        {
            RepairRates.PrintQuals(RepairRates.CountQuals(RepairRates.WeldersQuals(WeldersQualifications.GetQuals(), RepairRates.WelderNameFiller("Qual summary"))));
        }
    }
}
