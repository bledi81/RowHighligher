using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using System.Drawing;
using Office = Microsoft.Office.Core;

namespace RowHighligher
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        private ScientificCalculator calculator;

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
                this.ribbon.InvalidateControl("checkboxPlaceRuleOnTop");
                this.ribbon.InvalidateControl("checkboxMakeRuleBold");
                this.ribbon.InvalidateControl("checkboxCustomFontColor");
                this.ribbon.InvalidateControl("buttonChangeFontColor");
                this.ribbon.InvalidateControl("checkboxDetachCalculator");
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

        #region New UI Handlers

        public void OnPlaceRuleOnTop_Click(Office.IRibbonControl control, bool isPressed)
        {
            Properties.Settings.Default.PlaceRuleOnTop = isPressed;
            Properties.Settings.Default.Save();
            
            // Re-apply highlighting if enabled
            if (Globals.ThisAddIn.IsHighlighterEnabled)
            {
                Globals.ThisAddIn.ReapplyHighlighting();
            }
        }

        public bool GetPlaceRuleOnTop_Pressed(Office.IRibbonControl control)
        {
            return Properties.Settings.Default.PlaceRuleOnTop;
        }

        public void OnMakeRuleBold_Click(Office.IRibbonControl control, bool isPressed)
        {
            Properties.Settings.Default.MakeRuleBold = isPressed;
            Properties.Settings.Default.Save();
            
            // Re-apply highlighting if enabled
            if (Globals.ThisAddIn.IsHighlighterEnabled)
            {
                Globals.ThisAddIn.ReapplyHighlighting();
            }
        }

        public bool GetMakeRuleBold_Pressed(Office.IRibbonControl control)
        {
            return Properties.Settings.Default.MakeRuleBold;
        }

        public void OnCustomFontColor_Click(Office.IRibbonControl control, bool isPressed)
        {
            Properties.Settings.Default.CustomFontColorEnabled = isPressed;
            Properties.Settings.Default.Save();
            
            // Invalidate the font color button to update its enabled state
            if (this.ribbon != null)
            {
                this.ribbon.InvalidateControl("buttonChangeFontColor");
            }
            
            // Re-apply highlighting if enabled
            if (Globals.ThisAddIn.IsHighlighterEnabled)
            {
                Globals.ThisAddIn.ReapplyHighlighting();
            }
        }

        public bool GetCustomFontColor_Pressed(Office.IRibbonControl control)
        {
            return Properties.Settings.Default.CustomFontColorEnabled;
        }

        public void OnChangeFontColor_Click(Office.IRibbonControl control)
        {
            using (ColorDialog colorDialog = new ColorDialog())
            {
                colorDialog.Color = Properties.Settings.Default.CustomFontColor;
                colorDialog.AllowFullOpen = true;
                colorDialog.FullOpen = true;
                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    Properties.Settings.Default.CustomFontColor = colorDialog.Color;
                    Properties.Settings.Default.Save();

                    // Re-apply highlighting if enabled and using custom font color
                    if (Globals.ThisAddIn.IsHighlighterEnabled && Properties.Settings.Default.CustomFontColorEnabled)
                    {
                        Globals.ThisAddIn.ReapplyHighlighting();
                    }
                }
            }
        }

        public bool GetChangeFontColor_Enabled(Office.IRibbonControl control)
        {
            return Properties.Settings.Default.CustomFontColorEnabled;
        }

        public void OnShowCalculator_Click(Office.IRibbonControl control)
        {
            if (calculator == null || calculator.IsDisposed)
            {
                calculator = new ScientificCalculator();
                calculator.FormClosed += (s, e) => calculator = null;
                calculator.Show();
            }
            else
            {
                calculator.Activate();
            }
        }

        public void OnDetachCalculator_Click(Office.IRibbonControl control, bool isPressed)
        {
            Properties.Settings.Default.IsCalculatorDetached = isPressed;
            Properties.Settings.Default.Save();
            
            if (calculator != null && !calculator.IsDisposed)
            {
                calculator.TopMost = isPressed;
            }
        }

        public bool GetDetachCalculator_Pressed(Office.IRibbonControl control)
        {
            return Properties.Settings.Default.IsCalculatorDetached;
        }

        #endregion

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
