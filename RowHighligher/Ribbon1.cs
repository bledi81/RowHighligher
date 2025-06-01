using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;

namespace RowHighligher
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Ribbon1()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("RowHighligher.Ribbon1.xml");
        }

        #endregion

        #region Ribbon Callbacks

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
            if (this.ribbon != null)
            {
                this.ribbon.InvalidateControl("toggleEnableHighlighter"); 
            }
        }

        public void OnToggleEnableHighlighter_Click(Office.IRibbonControl control, bool isPressed)
        {
            Globals.ThisAddIn.IsHighlighterEnabled = isPressed;
            if (this.ribbon != null && control != null)
            {
                this.ribbon.InvalidateControl(control.Id);
            }
        }

        public bool GetEnableHighlighter_Pressed(Office.IRibbonControl control)
        {
            return Globals.ThisAddIn.IsHighlighterEnabled;
        }

        public void OnChangeColor_Click(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.ChangeHighlightColorWithDialog();
        }

        public void InvalidateToggleButton()
        {
            ribbon?.InvalidateControl("toggleEnableHighlighter"); 
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
