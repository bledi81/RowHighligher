using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Data;

namespace RowHighligher
{
    public partial class ThisAddIn
    {
        public Ribbon1 RibbonInstance { get; private set; }
        private const string ADDIN_CF_FORMULA = "=ROW()=ROW()+N(\"RowHighlighterAddInRule_v1.1\")";

        private bool wasHighlighterEnabledBeforeSave = false;

        // In the Calculator will use parenthesis to evaluate the expression correctly
        

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            if (Properties.Settings.Default.HighlightColor.A == 0)
            {
                Properties.Settings.Default.HighlightColor = System.Drawing.Color.Yellow;
                Properties.Settings.Default.Save();
            }
            
            if (Properties.Settings.Default.CustomFontColor.A == 0)
            {
                Properties.Settings.Default.CustomFontColor = System.Drawing.Color.Black;
                Properties.Settings.Default.Save();
            }

            this.Application.SheetSelectionChange += Application_SheetSelectionChange;
            this.Application.WorkbookBeforeSave += Application_WorkbookBeforeSave;
            this.Application.WorkbookBeforeClose += Application_WorkbookBeforeClose;
            this.Application.WorkbookAfterSave += Application_WorkbookAfterSave;

            Timer startupTimer = new Timer();
            startupTimer.Interval = 100;
            startupTimer.Tick += (s, args) => {
                startupTimer.Stop();
                startupTimer.Dispose();

                if (this.IsHighlighterEnabled)
                {
                    Excel.Range selection = null;
                    try
                    {
                        if (this.Application.ActiveWorkbook != null && this.Application.ActiveSheet != null)
                        {
                            selection = this.Application.Selection as Excel.Range;
                            if (selection != null)
                            {
                                System.Diagnostics.Debug.WriteLine("ThisAddIn_Startup (delayed): Applying initial highlighting.");
                                ApplyHighlightingToSelection(selection);
                            }
                        }
                    }
                    catch (COMException ex) { System.Diagnostics.Debug.WriteLine($"ThisAddIn_Startup (delayed) Error: {ex.Message}"); }
                    finally
                    {
                        if (selection != null) Marshal.ReleaseComObject(selection);
                    }
                }

                if (RibbonInstance != null)
                {
                    System.Diagnostics.Debug.WriteLine("ThisAddIn_Startup (delayed): Invalidating toggle button.");
                    RibbonInstance.InvalidateToggleButton();
                }
            };
            startupTimer.Start();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            if (this.Application != null)
            {
                this.Application.SheetSelectionChange -= Application_SheetSelectionChange;
                this.Application.WorkbookBeforeSave -= Application_WorkbookBeforeSave;
                this.Application.WorkbookBeforeClose -= Application_WorkbookBeforeClose;
                this.Application.WorkbookAfterSave += Application_WorkbookAfterSave;
            }

            try
            {
                if (this.Application != null && this.Application.Workbooks != null && this.Application.Workbooks.Count > 0)
                {
                    foreach (Excel.Workbook wb in this.Application.Workbooks)
                    {
                        foreach (Excel.Worksheet ws in wb.Worksheets)
                        {
                            RemoveAddinConditionalFormatting(ws);
                            Marshal.ReleaseComObject(ws);
                        }
                        Marshal.ReleaseComObject(wb);
                    }
                }
            }
            catch (COMException) { }
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            RibbonInstance = new Ribbon1();
            return RibbonInstance;
        }

        public bool IsHighlighterEnabled
        {
            get { return Properties.Settings.Default.IsHighlighterEnabled; }
            set
            {
                bool oldValue = Properties.Settings.Default.IsHighlighterEnabled;
                System.Diagnostics.Debug.WriteLine($"IsHighlighterEnabled_Set: OldValue={oldValue}, NewValue={value}");

                if (oldValue == value)
                {
                    if (value) {
                        System.Diagnostics.Debug.WriteLine("IsHighlighterEnabled_Set: Value is true and unchanged, re-applying to current selection.");
                        Excel.Range currentSelection = null;
                        try {
                            if (this.Application.ActiveWorkbook != null && this.Application.ActiveSheet != null) {
                                currentSelection = this.Application.Selection as Excel.Range;
                                if (currentSelection != null) {
                                    ApplyHighlightingToSelection(currentSelection);
                                }
                            }
                        } catch (COMException ex) { System.Diagnostics.Debug.WriteLine($"IsHighlighterEnabled_Set (no change, true) Error: {ex.Message}"); }
                        finally { if (currentSelection != null) Marshal.ReleaseComObject(currentSelection); }
                    }
                    RibbonInstance?.InvalidateToggleButton();
                    return;
                }

                Properties.Settings.Default.IsHighlighterEnabled = value;
                Properties.Settings.Default.Save();

                if (!value)
                {
                    System.Diagnostics.Debug.WriteLine("IsHighlighterEnabled_Set: Disabling. Removing all CF.");
                    RemoveAllAddinConditionalFormattingFromAllOpenWorksheets();
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("IsHighlighterEnabled_Set: Enabling. Applying to current selection.");
                    Excel.Range currentSelection = null;
                    try
                    {
                        if (this.Application.ActiveWorkbook != null && this.Application.ActiveSheet != null) {
                            currentSelection = this.Application.Selection as Excel.Range;
                            if (currentSelection != null)
                            {
                                ApplyHighlightingToSelection(currentSelection);
                            }
                        }
                    }
                    catch (COMException ex) { System.Diagnostics.Debug.WriteLine($"IsHighlighterEnabled_Set (changed to true) Error: {ex.Message}"); }
                    finally
                    {
                        if (currentSelection != null) Marshal.ReleaseComObject(currentSelection);
                    }
                }
                RibbonInstance?.InvalidateToggleButton();
            }
        }

        public void ChangeHighlightColorWithDialog()
        {
            using (ColorDialog colorDialog = new ColorDialog())
            {
                colorDialog.Color = Properties.Settings.Default.HighlightColor;
                colorDialog.AllowFullOpen = true;
                colorDialog.FullOpen = true;
                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    Properties.Settings.Default.HighlightColor = colorDialog.Color;
                    Properties.Settings.Default.Save();

                    if (this.IsHighlighterEnabled)
                    {
                        System.Diagnostics.Debug.WriteLine("ChangeHighlightColorWithDialog: Color changed, re-applying to current selection.");
                        Excel.Range selection = null;
                        try
                        {
                             if (this.Application.ActiveWorkbook != null && this.Application.ActiveSheet != null) {
                                selection = this.Application.Selection as Excel.Range;
                                if (selection != null)
                                {
                                    ApplyHighlightingToSelection(selection);
                                }
                             }
                        }
                        catch (COMException ex) { System.Diagnostics.Debug.WriteLine($"ChangeHighlightColorWithDialog Error: {ex.Message}"); }
                        finally
                        {
                            if (selection != null) Marshal.ReleaseComObject(selection);
                        }
                    }
                }
            }
        }
        
        // This method can be called when settings change to reapply highlighting to the current selection
        public void ReapplyHighlighting()
        {
            if (this.IsHighlighterEnabled)
            {
                System.Diagnostics.Debug.WriteLine("ReapplyHighlighting: Re-applying to current selection.");
                Excel.Range selection = null;
                try
                {
                    if (this.Application.ActiveWorkbook != null && this.Application.ActiveSheet != null) {
                        selection = this.Application.Selection as Excel.Range;
                        if (selection != null)
                        {
                            ApplyHighlightingToSelection(selection);
                        }
                    }
                }
                catch (COMException ex) { System.Diagnostics.Debug.WriteLine($"ReapplyHighlighting Error: {ex.Message}"); }
                finally
                {
                    if (selection != null) Marshal.ReleaseComObject(selection);
                }
            }
        }

        private void Application_SheetSelectionChange(object Sh, Excel.Range Target)
        {
            System.Diagnostics.Debug.WriteLine($"SheetSelectionChange: Fired. IsHighlighterEnabled: {this.IsHighlighterEnabled}. Event Target: {Target?.Address ?? "null"}");
            Excel.Range selectionToHighlight = null;

            if (this.IsHighlighterEnabled)
            {
                try
                {
                    if (this.Application.ActiveWorkbook != null && this.Application.ActiveSheet != null) {
                        selectionToHighlight = this.Application.Selection as Excel.Range;
                        if (selectionToHighlight != null)
                        {
                            System.Diagnostics.Debug.WriteLine($"SheetSelectionChange: Using fresh selection from App: {selectionToHighlight.Address ?? "null"}");
                        } else if (Target != null) {
                             System.Diagnostics.Debug.WriteLine($"SheetSelectionChange: Fresh selection was null, falling back to Event Target: {Target.Address ?? "null"}");
                            selectionToHighlight = Target;
                        }
                    } else if (Target != null) {
                        System.Diagnostics.Debug.WriteLine($"SheetSelectionChange: No active workbook/sheet, falling back to Event Target: {Target.Address ?? "null"}");
                        selectionToHighlight = Target;
                    }

                    if (selectionToHighlight != null)
                    {
                        ApplyHighlightingToSelection(selectionToHighlight);
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine("SheetSelectionChange: No valid selection to highlight.");
                    }
                }
                catch (COMException ex)
                {
                    System.Diagnostics.Debug.WriteLine($"SheetSelectionChange: Error during processing: {ex.Message}");
                }
                finally
                {
                    if (selectionToHighlight != null && selectionToHighlight != Target)
                    {
                        Marshal.ReleaseComObject(selectionToHighlight);
                    }
                }
            }

            if (Target != null) Marshal.ReleaseComObject(Target);
            if (Sh != null) Marshal.ReleaseComObject(Sh);
        }

        private void Application_WorkbookBeforeSave(Excel.Workbook Wb, bool SaveAsUI, ref bool Cancel)
        {
            System.Diagnostics.Debug.WriteLine("WorkbookBeforeSave: Storing highlighter state before save.");

            // Remember the current state before disabling
            wasHighlighterEnabledBeforeSave = this.IsHighlighterEnabled;

            // Disable highlighting during save
            if (wasHighlighterEnabledBeforeSave)
            {
                System.Diagnostics.Debug.WriteLine("WorkbookBeforeSave: Temporarily disabling highlighter for save.");
                this.IsHighlighterEnabled = false;
            }
        }

        private void Application_WorkbookBeforeClose(Excel.Workbook Wb, ref bool Cancel)
        {
            System.Diagnostics.Debug.WriteLine("WorkbookBeforeClose: Toggling off highlighter before close.");
            this.IsHighlighterEnabled = false;
        }

        private void Application_WorkbookAfterSave(Excel.Workbook Wb, bool Success)
        {
            System.Diagnostics.Debug.WriteLine($"WorkbookAfterSave: Restoring highlighter to previous state: {wasHighlighterEnabledBeforeSave}");

            // Only re-enable if it was enabled before saving
            if (wasHighlighterEnabledBeforeSave)
            {
                this.IsHighlighterEnabled = true;
            }

        }

        private void ApplyHighlightingToSelection(Excel.Range selection)
        {
            System.Diagnostics.Debug.WriteLine($"ApplyHighlightingToSelection: START. Selection Address: '{selection?.Address ?? "null"}'. IsHighlighterEnabled: {this.IsHighlighterEnabled}");
            if (selection == null || !this.IsHighlighterEnabled)
            {
                System.Diagnostics.Debug.WriteLine("ApplyHighlightingToSelection: Exiting early (selection is null or highlighter disabled).");
                return;
            }

            Excel.Worksheet activeSheet = null;
            Excel.Areas areas = null;
            Excel.Range entireRow = null;
            Excel.Areas entireRowAreas = null;
            Excel.Range rowSubArea = null;
            Excel.FormatConditions fcs = null;
            Excel.FormatCondition fc = null;
            
            List<Excel.Range> rangesFromMap = new List<Excel.Range>();

            try
            {
                activeSheet = selection.Worksheet;
                System.Diagnostics.Debug.WriteLine($"ApplyHighlightingToSelection: Processing sheet '{activeSheet.Name}'. Attempting to remove existing CF.");
                RemoveAddinConditionalFormatting(activeSheet);
                System.Diagnostics.Debug.WriteLine($"ApplyHighlightingToSelection: Finished removing CF from sheet '{activeSheet.Name}'.");

                Dictionary<string, Excel.Range> uniqueRowAddressToRangeMap = new Dictionary<string, Excel.Range>();
                areas = selection.Areas;
                System.Diagnostics.Debug.WriteLine($"ApplyHighlightingToSelection: Selection has {areas.Count} area(s).");

                foreach (Excel.Range area in areas)
                {
                    entireRow = area.EntireRow;
                    entireRowAreas = entireRow.Areas;
                    foreach(Excel.Range currentRSB in entireRowAreas)
                    {
                        rowSubArea = currentRSB;
                         if (!uniqueRowAddressToRangeMap.ContainsKey(rowSubArea.Address))
                         {
                             uniqueRowAddressToRangeMap.Add(rowSubArea.Address, rowSubArea);
                             rangesFromMap.Add(rowSubArea);
                         }
                         else 
                         {
                             Marshal.ReleaseComObject(rowSubArea);
                             rowSubArea = null;
                         }
                    }
                    if (entireRowAreas != null) { Marshal.ReleaseComObject(entireRowAreas); entireRowAreas = null; }
                    if (entireRow != null) { Marshal.ReleaseComObject(entireRow); entireRow = null; }
                    if (area != null) Marshal.ReleaseComObject(area);
                }
                if (areas != null) { Marshal.ReleaseComObject(areas); areas = null; }

                System.Diagnostics.Debug.WriteLine($"ApplyHighlightingToSelection: Found {uniqueRowAddressToRangeMap.Count} unique row areas to format.");
                foreach (Excel.Range rowRangeToFormat in uniqueRowAddressToRangeMap.Values)
                {
                    System.Diagnostics.Debug.WriteLine($"ApplyHighlightingToSelection: Adding CF to {rowRangeToFormat.Address}");
                    fcs = rowRangeToFormat.FormatConditions;
                    
                    // Use the priority parameter to set the rule on top if requested
                    int priority = Properties.Settings.Default.PlaceRuleOnTop ? 1 : fcs.Count + 1;
                    
                    fc = (Excel.FormatCondition)fcs.Add(
                        Excel.XlFormatConditionType.xlExpression,
                        Formula1: ADDIN_CF_FORMULA);
                    
                    // Set the rule's priority (position in the stack)
                    fc.Priority = priority;
                    
                    // Apply background color
                    fc.Interior.Color = ColorTranslator.ToOle(Properties.Settings.Default.HighlightColor);
                    
                    // Apply font bold if enabled
                    if (Properties.Settings.Default.MakeRuleBold)
                    {
                        fc.Font.Bold = true;
                    }
                    
                    // Apply font color if enabled
                    if (Properties.Settings.Default.CustomFontColorEnabled)
                    {
                        fc.Font.Color = ColorTranslator.ToOle(Properties.Settings.Default.CustomFontColor);
                    }
                    
                    fc.StopIfTrue = false;
                    
                    if (fc != null) { Marshal.ReleaseComObject(fc); fc = null; }
                    if (fcs != null) { Marshal.ReleaseComObject(fcs); fcs = null; }
                }
                System.Diagnostics.Debug.WriteLine($"ApplyHighlightingToSelection: Successfully applied new CF rules.");
            }
            catch (COMException ex)
            {
                System.Diagnostics.Debug.WriteLine($"ApplyHighlightingToSelection: ERROR - {ex.Message} (StackTrace: {ex.StackTrace})");
            }
            finally
            {
                if (fc != null) Marshal.ReleaseComObject(fc);
                if (fcs != null) Marshal.ReleaseComObject(fcs);
                if (rowSubArea != null) Marshal.ReleaseComObject(rowSubArea);
                if (entireRowAreas != null) Marshal.ReleaseComObject(entireRowAreas);
                if (entireRow != null) Marshal.ReleaseComObject(entireRow);
                if (areas != null) Marshal.ReleaseComObject(areas);
                foreach(var r in rangesFromMap) Marshal.ReleaseComObject(r);
                if (activeSheet != null) Marshal.ReleaseComObject(activeSheet);
            }
            System.Diagnostics.Debug.WriteLine($"ApplyHighlightingToSelection: END. Selection Address: '{selection?.Address ?? "null"}'.");
        }

        private void RemoveAddinConditionalFormatting(Excel.Worksheet ws)
        {
            System.Diagnostics.Debug.WriteLine($"RemoveAddinConditionalFormatting: START for sheet '{ws?.Name ?? "null"}'");
            if (ws == null) return;

            Excel.FormatConditions fcs = null;
            Excel.FormatCondition fc = null;
            List<Excel.FormatCondition> toDelete = new List<Excel.FormatCondition>();
            int initialCount = 0;

            try
            {
                fcs = ws.Cells.FormatConditions;
                initialCount = fcs.Count;
                if (initialCount == 0) {
                    System.Diagnostics.Debug.WriteLine($"RemoveAddinConditionalFormatting: No CF rules on sheet '{ws.Name}'.");
                    if (fcs != null) Marshal.ReleaseComObject(fcs);
                    return;
                }
                System.Diagnostics.Debug.WriteLine($"RemoveAddinConditionalFormatting: Sheet '{ws.Name}' has {initialCount} CF rules. Checking for add-in rules.");

                for (int i = initialCount; i >= 1; i--)
                {
                    fc = fcs[i];
                    if (fc.Type == (int)Excel.XlFormatConditionType.xlExpression &&
                        fc.Formula1 == ADDIN_CF_FORMULA)
                    {
                        toDelete.Add(fc); 
                    }
                    else
                    {
                        Marshal.ReleaseComObject(fc);
                    }
                    fc = null;
                }

                System.Diagnostics.Debug.WriteLine($"RemoveAddinConditionalFormatting: Found {toDelete.Count} add-in rules to delete on sheet '{ws.Name}'.");
                foreach (Excel.FormatCondition fcToDelete in toDelete)
                {
                    fcToDelete.Delete();
                    Marshal.ReleaseComObject(fcToDelete);
                }
            }
            catch (COMException ex)
            {
                System.Diagnostics.Debug.WriteLine($"RemoveAddinConditionalFormatting: ERROR - {ex.Message}");
            }
            finally
            {
                if (fc != null) Marshal.ReleaseComObject(fc);
                if (fcs != null) Marshal.ReleaseComObject(fcs);
                System.Diagnostics.Debug.WriteLine($"RemoveAddinConditionalFormatting: END for sheet '{ws?.Name ?? "null"}'");
            }
        }

        private void RemoveAllAddinConditionalFormattingFromAllOpenWorksheets()
        {
            System.Diagnostics.Debug.WriteLine("RemoveAllAddinConditionalFormattingFromAllOpenWorksheets: START");
            Excel.Workbooks wbs = null;
            Excel.Workbook wb = null;
            Excel.Sheets wss = null; // Changed from Excel.Worksheets to Excel.Sheets for compatibility with wb.Worksheets
            Excel.Worksheet ws = null;
            try
            {
                if (this.Application != null) {
                    wbs = this.Application.Workbooks;
                    if (wbs != null && wbs.Count > 0)
                    {
                        for(int i=1; i <= wbs.Count; i++) // COM collections are 1-indexed
                        {
                            wb = wbs[i];
                            wss = wb.Worksheets; // Corrected: Use Worksheets property to get all sheets
                            if (wss != null) // Add null check for wss
                            {
                                foreach (object sheetObj in wss) // Iterate using foreach for safety
                                {
                                    if (sheetObj is Excel.Worksheet)
                                    {
                                        ws = (Excel.Worksheet)sheetObj;
                                        RemoveAddinConditionalFormatting(ws);
                                        if (ws != null) { Marshal.ReleaseComObject(ws); ws = null; }
                                    }
                                    else // Release non-worksheet objects if any (e.g. Charts)
                                    {
                                        Marshal.ReleaseComObject(sheetObj);
                                    }
                                }
                                if (wss != null) { Marshal.ReleaseComObject(wss); wss = null; }
                            }
                            if (wb != null) { Marshal.ReleaseComObject(wb); wb = null; }
                        }
                    }
                    if (wbs != null) { Marshal.ReleaseComObject(wbs); wbs = null; }
                }
            }
            catch (COMException ex) { System.Diagnostics.Debug.WriteLine($"RemoveAllAddinConditionalFormattingFromAllOpenWorksheets: ERROR - {ex.Message}"); }
            finally {
                // Ensure all COM objects in this scope are released
                if (ws != null) Marshal.ReleaseComObject(ws); // ws should be null if released in loop
                if (wss != null) Marshal.ReleaseComObject(wss); // wss should be null if released in loop
                if (wb != null) Marshal.ReleaseComObject(wb);   // wb should be null if released in loop
                if (wbs != null) Marshal.ReleaseComObject(wbs);
                System.Diagnostics.Debug.WriteLine("RemoveAllAddinConditionalFormattingFromAllOpenWorksheets: END");
            }
        }

        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

    }
}