using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;

namespace CosmeticSpecAddinCS
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Globals.ThisAddIn.Application.Documents.Add(@"\\storage03.se.axis.com\hw-apps\ptc\part_specification_deviation\Template\PSDTemplate.dotm");
                Globals.Ribbons.Ribbon1.RibbonUI.ActivateTab("tab1");
            }
            catch (Exception)
            {

                throw;
            }
        }

        private void button7_Click(object sender, RibbonControlEventArgs e)
        {
            var box = new AboutBox1();
            box.Show();
        }

        private void button8_Click(object sender, RibbonControlEventArgs e)
        {
            System.Diagnostics.Process.Start("https://glinky.axis.com/mek_psd");

        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Globals.ThisAddIn.Application.Run("SaveDeviation");

            }
            catch (Exception)
            {
                MessageBox.Show("This is not a Part Specification Deviation!");
            }
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Globals.ThisAddIn.Application.Run("insertblock");
            }
            catch (Exception)
            {

                throw;
            }
        }
    }
}
