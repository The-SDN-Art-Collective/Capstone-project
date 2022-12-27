using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Visio = Microsoft.Office.Interop.Visio;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using Microsoft.Office.Interop.Visio;

namespace VisioAddIn1
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {

        }

        /// <summary>
        /// Draws a Rectangle, returns the shape
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <returns></returns>
        public Visio.Shape DrawRectangle( double x, double y)
        {

           Visio.Documents visioDocs = this.Application.Documents;
            Visio.Document visioStencil = visioDocs.OpenEx("Basic Shapes.vss",
                (short)Microsoft.Office.Interop.Visio.VisOpenSaveArgs.visOpenDocked);

            Visio.Page visioPage = this.Application.ActivePage;

            Visio.Master visioRectMaster = visioStencil.Masters.get_ItemU(@"Cross");
            Visio.Shape visioRectShape = visioPage.Drop(visioRectMaster, x, y);
            visioRectShape.Text = @"Rectangle text.";

            return visioRectShape;
            
      
        }

        /// <summary>
        /// Draws a Hexagon, returns the shape
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <returns></returns>
        public Visio.Shape DrawHex(double x, double y)
        {

            Visio.Documents visioDocs = this.Application.Documents;
            Visio.Document visioStencil = visioDocs.OpenEx("Basic Shapes.vss",
                (short)Microsoft.Office.Interop.Visio.VisOpenSaveArgs.visOpenDocked);

            Visio.Page visioPage = this.Application.ActivePage;

            Visio.Master visioHexagonMaster = visioStencil.Masters.get_ItemU(@"Hexagon");
            Visio.Shape visioHexagonShape = visioPage.Drop(visioHexagonMaster, x, y);
            visioHexagonShape.Text = @"Hexagon text.";

            return visioHexagonShape;
        }

        public void DrawConnection(double x, double y)
        {
           Visio.Shape Rectangle = DrawRectangle(x, y);
           Visio.Shape Hexagon = DrawHex(x+5, y-5);

            Rectangle.AutoConnect(Hexagon, VisAutoConnectDir.visAutoConnectDirNone);

            
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
