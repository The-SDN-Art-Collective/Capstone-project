using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;




namespace VisioAddIn1
{

    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            //isioAddIn1.ThisAddIn MainOb = Globals.ThisAddIn;
            Globals.ThisAddIn.DrawRectangle();
=
        }

        private void button1_Click_1(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.DrawHex();

        }
    }
}
