using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Data;
using System.Threading;
using System.IO;
using System.Diagnostics;

namespace RowHighligher
{
    public partial class ThisAddIn
    {
        public Ribbon1 RibbonInstance { get; private set; }
        private const string ADDIN_CF_FORMULA = "=ROW()=ROW()+N(\"RowHighlighterAddInRule_v1.1\")";

        private bool wasHighlighterEnabledBeforeSave = false;
        private bool globalHighlighterState = false;

        // Flag to track state for event handling safety
        private bool isProcessingOperation = false;
        private bool isRightClickInProgress = false;

        // Clipboard helpers
        private object savedClipboardData = null;
        private string savedClipboardFormat = null;
        private DateTime lastClipboardSaveTime = DateTime.MinValue;

        // Clipboard operations interop
        [DllImport("user32.dll")]
        private static extern IntPtr GetOpenClipboardWindow();

        [DllImport("user32.dll")]
        private static extern bool OpenClipboard(IntPtr hWndNewOwner);

        [DllImport("user32.dll")]
        private static extern bool CloseClipboard();

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

            // Register Excel event handlers
            this.Application.SheetSelectionChange += Application_SheetSelectionChange;
            this.Application.SheetBeforeRightClick += Application_SheetBeforeRightClick;
            this.Application.WorkbookBeforeSave += Application_WorkbookBeforeSave;
            this.Application.WorkbookBeforeClose += Application_WorkbookBeforeClose;
            this.Application.WorkbookAfterSave += Application_WorkbookAfterSave;
            this.Application.WorkbookOpen += Application_WorkbookOpen;

            // Use a delayed initialization to allow Excel to fully load
            System.Windows.Forms.Timer startupTimer = new System.Windows.Forms.Timer();
            startupTimer.Interval = 100;
            startupTimer.Tick += (s, args) =>
            {
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
                                DoApplyHighlightingToSelection(selection);
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
                // Unregister all event handlers
                this.Application.SheetSelectionChange -= Application_SheetSelectionChange;
                this.Application.SheetBeforeRightClick -= Application_SheetBeforeRightClick;
                this.Application.WorkbookBeforeSave -= Application_WorkbookBeforeSave;
                this.Application.WorkbookBeforeClose -= Application_WorkbookBeforeClose;
                this.Application.WorkbookAfterSave -= Application_WorkbookAfterSave;
                this.Application.WorkbookOpen -= Application_WorkbookOpen;
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

        // Handle right-clicks to prevent clipboard and context menu issues
        private void Application_SheetBeforeRightClick(object Sh, Excel.Range Target, ref bool Cancel)
        {
            if (isProcessingOperation)
            {
                // If we're already processing something, let the right-click proceed normally
                System.Diagnostics.Debug.WriteLine("SheetBeforeRightClick: Already processing, allowing right-click");
                return;
            }

            try
            {
                isRightClickInProgress = true;

                // Save current clipboard data before proceeding
                SaveClipboardContentsIfNeeded();

                // Let the right-click proceed normally
                System.Diagnostics.Debug.WriteLine("SheetBeforeRightClick: Allowing right-click with clipboard preserved");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"SheetBeforeRightClick Error: {ex.Message}");
            }
            finally
            {
                // Release COM objects correctly
                if (Target != null) Marshal.ReleaseComObject(Target);
                if (Sh != null) Marshal.ReleaseComObject(Sh);
            }
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
                    if (value)
                    {
                        System.Diagnostics.Debug.WriteLine("IsHighlighterEnabled_Set: Value is true and unchanged, re-applying to current selection.");
                        Excel.Range currentSelection = null;
                        try
                        {
                            if (this.Application.ActiveWorkbook != null && this.Application.ActiveSheet != null)
                            {
                                currentSelection = this.Application.Selection as Excel.Range;
                                if (currentSelection != null)
                                {
                                    // Use our utility method for applying highlighting
                                    DoApplyHighlightingToSelection(currentSelection);
                                }
                            }
                        }
                        catch (COMException ex) { System.Diagnostics.Debug.WriteLine($"IsHighlighterEnabled_Set (no change, true) Error: {ex.Message}"); }
                        finally { if (currentSelection != null) Marshal.ReleaseComObject(currentSelection); }
                    }
                    RibbonInstance?.InvalidateToggleButton();
                    return;
                }

                // Save the current clipboard data
                SaveClipboardContentsIfNeeded();

                try
                {
                    isProcessingOperation = true;

                    Properties.Settings.Default.IsHighlighterEnabled = value;
                    Properties.Settings.Default.Save();
                    globalHighlighterState = value; // Keep global state in sync

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
                            if (this.Application.ActiveWorkbook != null && this.Application.ActiveSheet != null)
                            {
                                currentSelection = this.Application.Selection as Excel.Range;
                                if (currentSelection != null)
                                {
                                    DoApplyHighlightingToSelection(currentSelection);
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
                finally
                {
                    isProcessingOperation = false;

                    // Restore clipboard at the end
                    RestoreClipboardData();
                }
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
                    // Save current clipboard state
                    SaveClipboardContentsIfNeeded();

                    try
                    {
                        isProcessingOperation = true;

                        Properties.Settings.Default.HighlightColor = colorDialog.Color;
                        Properties.Settings.Default.Save();

                        if (this.IsHighlighterEnabled)
                        {
                            System.Diagnostics.Debug.WriteLine("ChangeHighlightColorWithDialog: Color changed, re-applying to current selection.");
                            Excel.Range selection = null;
                            try
                            {
                                if (this.Application.ActiveWorkbook != null && this.Application.ActiveSheet != null)
                                {
                                    selection = this.Application.Selection as Excel.Range;
                                    if (selection != null)
                                    {
                                        DoApplyHighlightingToSelection(selection);
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
                    finally
                    {
                        isProcessingOperation = false;

                        // Restore clipboard data
                        RestoreClipboardData();
                    }
                }
            }
        }

        public void ReapplyHighlighting()
        {
            if (this.IsHighlighterEnabled)
            {
                // Save current clipboard state
                SaveClipboardContentsIfNeeded();

                try
                {
                    isProcessingOperation = true;

                    System.Diagnostics.Debug.WriteLine("ReapplyHighlighting: Re-applying to current selection.");
                    Excel.Range selection = null;
                    try
                    {
                        if (this.Application.ActiveWorkbook != null && this.Application.ActiveSheet != null)
                        {
                            selection = this.Application.Selection as Excel.Range;
                            if (selection != null)
                            {
                                DoApplyHighlightingToSelection(selection);
                            }
                        }
                    }
                    catch (COMException ex) { System.Diagnostics.Debug.WriteLine($"ReapplyHighlighting Error: {ex.Message}"); }
                    finally
                    {
                        if (selection != null) Marshal.ReleaseComObject(selection);
                    }
                }
                finally
                {
                    isProcessingOperation = false;

                    // Restore clipboard data
                    RestoreClipboardData();
                }
            }
        }

        // This is the main event handler that gets called when the selection changes in Excel
        private void Application_SheetSelectionChange(object Sh, Excel.Range Target)
        {
            // For right-click operations, we prefer to handle them differently
            if (isRightClickInProgress)
            {
                System.Diagnostics.Debug.WriteLine("SheetSelectionChange: Right-click in progress, deferring highlight");
                isRightClickInProgress = false;
                if (Target != null) Marshal.ReleaseComObject(Target);
                if (Sh != null) Marshal.ReleaseComObject(Sh);
                return;
            }

            // If already processing an operation, skip to avoid recursion
            if (isProcessingOperation)
            {
                System.Diagnostics.Debug.WriteLine("SheetSelectionChange: Already processing, skipping");
                if (Target != null) Marshal.ReleaseComObject(Target);
                if (Sh != null) Marshal.ReleaseComObject(Sh);
                return;
            }

            System.Diagnostics.Debug.WriteLine($"SheetSelectionChange: Fired. IsHighlighterEnabled: {this.IsHighlighterEnabled}. Event Target: {Target?.Address ?? "null"}");
            Excel.Range selectionToHighlight = null;

            if (this.IsHighlighterEnabled)
            {
                // Save clipboard content before making any changes
                SaveClipboardContentsIfNeeded();

                try
                {
                    isProcessingOperation = true;

                    if (this.Application.ActiveWorkbook != null && this.Application.ActiveSheet != null)
                    {
                        selectionToHighlight = this.Application.Selection as Excel.Range;
                        if (selectionToHighlight != null)
                        {
                            System.Diagnostics.Debug.WriteLine($"SheetSelectionChange: Using fresh selection from App: {selectionToHighlight.Address ?? "null"}");
                        }
                        else if (Target != null)
                        {
                            System.Diagnostics.Debug.WriteLine($"SheetSelectionChange: Fresh selection was null, falling back to Event Target: {Target.Address ?? "null"}");
                            selectionToHighlight = Target;
                        }
                    }
                    else if (Target != null)
                    {
                        System.Diagnostics.Debug.WriteLine($"SheetSelectionChange: No active workbook/sheet, falling back to Event Target: {Target.Address ?? "null"}");
                        selectionToHighlight = Target;
                    }

                    if (selectionToHighlight != null)
                    {
                        DoApplyHighlightingToSelection(selectionToHighlight);
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

                    isProcessingOperation = false;

                    // Restore clipboard content after we're done
                    RestoreClipboardData();
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
            System.Diagnostics.Debug.WriteLine("WorkbookBeforeClose: Removing highlighter from closing workbook.");
            
            // Close calculator form if it's open
            if (RibbonInstance?.calculator != null && !RibbonInstance.calculator.IsDisposed)
            {
                System.Diagnostics.Debug.WriteLine("WorkbookBeforeClose: Closing calculator form");
                RibbonInstance.calculator.Close();
                // No need to set RibbonInstance.calculator = null as FormClosed event handler does this
            }
            
            // Close units converter form if it's open
            if (RibbonInstance?.unitsConverter != null && !RibbonInstance.unitsConverter.IsDisposed)
            {
                System.Diagnostics.Debug.WriteLine("WorkbookBeforeClose: Closing units converter form");
                RibbonInstance.unitsConverter.Close();
                // No need to set RibbonInstance.unitsConverter = null as FormClosed event handler does this
            }
            
            // Original code to remove highlighting from the workbook being closed
            if (this.IsHighlighterEnabled)
            {
                foreach (Excel.Worksheet ws in Wb.Worksheets)
                {
                    RemoveAddinConditionalFormatting(ws);
                    Marshal.ReleaseComObject(ws);
                }
            }
            // Do NOT change Properties.Settings.Default.IsHighlighterEnabled here.
            // Do NOT save settings here.
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

        private void Application_WorkbookOpen(Excel.Workbook Wb)
        {
            System.Diagnostics.Debug.WriteLine($"WorkbookOpen: Highlighter state is {this.IsHighlighterEnabled}");
            // The IsHighlighterEnabled state should already be loaded correctly from settings
            // during ThisAddIn_Startup. If it's enabled, apply to the newly opened workbook.
            if (this.IsHighlighterEnabled)
            {
                Excel.Worksheet activeSheet = null;
                Excel.Range selection = null;
                try
                {
                    activeSheet = Wb.ActiveSheet as Excel.Worksheet;
                    if (activeSheet != null)
                    {
                        selection = this.Application.Selection as Excel.Range; // Or Wb.Application.Selection
                        if (selection != null && selection.Worksheet.Parent.Name == Wb.Name) // Ensure selection is in the opened workbook
                        {
                            DoApplyHighlightingToSelection(selection);
                        }
                        else if (selection != null)
                        {
                            Marshal.ReleaseComObject(selection); // Release if it's from another workbook
                        }
                    }
                }
                catch (COMException ex)
                {
                    System.Diagnostics.Debug.WriteLine($"WorkbookOpen ApplyHighlighting Error: {ex.Message}");
                }
                finally
                {
                    if (selection != null) Marshal.ReleaseComObject(selection);
                    if (activeSheet != null) Marshal.ReleaseComObject(activeSheet);
                }
            }
            // Ribbon should be updated in ThisAddIn_Startup or when IsHighlighterEnabled is changed.
            // RibbonInstance?.InvalidateToggleButton(); // May not be needed here if startup handles it.
        }

        // Private helper methods
        private void SaveClipboardContentsIfNeeded()
        {
            // Don't re-save within a short period to avoid excessive operations
            if ((DateTime.Now - lastClipboardSaveTime).TotalMilliseconds < 500)
            {
                return;
            }

            // Only save if we don't already have the clipboard data
            if (savedClipboardData == null)
            {
                try
                {
                    // Check if clipboard is open by another application
                    IntPtr clipboardWindow = GetOpenClipboardWindow();
                    if (clipboardWindow != IntPtr.Zero && clipboardWindow != Process.GetCurrentProcess().MainWindowHandle)
                    {
                        // Skip saving if the clipboard is being used by another process
                        return;
                    }

                    // Save clipboard data based on available formats
                    if (Clipboard.ContainsText())
                    {
                        savedClipboardData = Clipboard.GetText();
                        savedClipboardFormat = DataFormats.Text;
                        System.Diagnostics.Debug.WriteLine("Saved clipboard text data");
                    }
                    else if (Clipboard.ContainsImage())
                    {
                        savedClipboardData = Clipboard.GetImage();
                        savedClipboardFormat = DataFormats.Bitmap;
                        System.Diagnostics.Debug.WriteLine("Saved clipboard image data");
                    }
                    else if (Clipboard.ContainsFileDropList())
                    {
                        savedClipboardData = Clipboard.GetFileDropList();
                        savedClipboardFormat = DataFormats.FileDrop;
                        System.Diagnostics.Debug.WriteLine("Saved clipboard file list");
                    }
                    else if (Clipboard.ContainsData(DataFormats.Html))
                    {
                        savedClipboardData = Clipboard.GetData(DataFormats.Html);
                        savedClipboardFormat = DataFormats.Html;
                        System.Diagnostics.Debug.WriteLine("Saved clipboard HTML data");
                    }
                    else if (Clipboard.ContainsData(DataFormats.Rtf))
                    {
                        savedClipboardData = Clipboard.GetData(DataFormats.Rtf);
                        savedClipboardFormat = DataFormats.Rtf;
                        System.Diagnostics.Debug.WriteLine("Saved clipboard RTF data");
                    }
                    else if (Clipboard.ContainsAudio())
                    {
                        savedClipboardData = Clipboard.GetAudioStream();
                        savedClipboardFormat = DataFormats.WaveAudio;
                        System.Diagnostics.Debug.WriteLine("Saved clipboard audio data");
                    }

                    // Update the timestamp even if we didn't save anything
                    lastClipboardSaveTime = DateTime.Now;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error saving clipboard: {ex.Message}");
                    savedClipboardData = null;
                    savedClipboardFormat = null;
                }
            }
        }

        private void RestoreClipboardData()
        {
            // Only restore if we actually have data
            if (savedClipboardData != null && !string.IsNullOrEmpty(savedClipboardFormat))
            {
                try
                {
                    // Check if clipboard is open by another application
                    IntPtr clipboardWindow = GetOpenClipboardWindow();
                    if (clipboardWindow != IntPtr.Zero && clipboardWindow != Process.GetCurrentProcess().MainWindowHandle)
                    {
                        // Skip restoring if the clipboard is being used by another process
                        return;
                    }

                    // We need a brief delay to make sure clipboard operations finish
                    Thread.Sleep(50);

                    // Restore based on saved format
                    if (savedClipboardFormat == DataFormats.Text && savedClipboardData is string textData)
                    {
                        Clipboard.SetText(textData);
                        System.Diagnostics.Debug.WriteLine("Restored clipboard text data");
                    }
                    else if (savedClipboardFormat == DataFormats.Bitmap && savedClipboardData is Image imageData)
                    {
                        Clipboard.SetImage(imageData);
                        System.Diagnostics.Debug.WriteLine("Restored clipboard image data");
                    }
                    else if (savedClipboardFormat == DataFormats.FileDrop && savedClipboardData is System.Collections.Specialized.StringCollection fileList)
                    {
                        Clipboard.SetFileDropList(fileList);
                        System.Diagnostics.Debug.WriteLine("Restored clipboard file list");
                    }
                    else if ((savedClipboardFormat == DataFormats.Html ||
                              savedClipboardFormat == DataFormats.Rtf) &&
                              savedClipboardData != null)
                    {
                        Clipboard.SetData(savedClipboardFormat, savedClipboardData);
                        System.Diagnostics.Debug.WriteLine($"Restored clipboard {savedClipboardFormat} data");
                    }
                    else if (savedClipboardFormat == DataFormats.WaveAudio && savedClipboardData is Stream audioData)
                    {
                        audioData.Position = 0;
                        Clipboard.SetAudio(audioData);
                        System.Diagnostics.Debug.WriteLine("Restored clipboard audio data");
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error restoring clipboard: {ex.Message}");
                }
                finally
                {
                    // Clear the saved data after attempting to restore
                    savedClipboardData = null;
                    savedClipboardFormat = null;
                }
            }
        }

        // Implementation of the core highlighting functionality
        private void DoApplyHighlightingToSelection(Excel.Range selection)
        {
            System.Diagnostics.Debug.WriteLine($"DoApplyHighlightingToSelection: START. Selection Address: '{selection?.Address ?? "null"}'.");
            if (selection == null || !this.IsHighlighterEnabled)
            {
                System.Diagnostics.Debug.WriteLine("DoApplyHighlightingToSelection: Exiting early (selection is null or highlighter disabled).");
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
                System.Diagnostics.Debug.WriteLine($"DoApplyHighlightingToSelection: Processing sheet '{activeSheet.Name}'.");
                RemoveAddinConditionalFormatting(activeSheet);

                Dictionary<string, Excel.Range> uniqueRowAddressToRangeMap = new Dictionary<string, Excel.Range>();
                areas = selection.Areas;
                System.Diagnostics.Debug.WriteLine($"DoApplyHighlightingToSelection: Selection has {areas.Count} area(s).");

                foreach (Excel.Range area in areas)
                {
                    entireRow = area.EntireRow;
                    entireRowAreas = entireRow.Areas;
                    foreach (Excel.Range currentRSB in entireRowAreas)
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

                System.Diagnostics.Debug.WriteLine($"DoApplyHighlightingToSelection: Found {uniqueRowAddressToRangeMap.Count} unique row areas to format.");

                // Add conditional formatting one by one to each range
                foreach (Excel.Range rowRangeToFormat in uniqueRowAddressToRangeMap.Values)
                {
                    System.Diagnostics.Debug.WriteLine($"DoApplyHighlightingToSelection: Adding CF to {rowRangeToFormat.Address}");
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
                System.Diagnostics.Debug.WriteLine($"DoApplyHighlightingToSelection: Successfully applied new CF rules.");
            }
            catch (COMException ex)
            {
                System.Diagnostics.Debug.WriteLine($"DoApplyHighlightingToSelection: ERROR - {ex.Message}");
            }
            finally
            {
                // Release all COM objects properly
                if (fc != null) Marshal.ReleaseComObject(fc);
                if (fcs != null) Marshal.ReleaseComObject(fcs);
                if (rowSubArea != null) Marshal.ReleaseComObject(rowSubArea);
                if (entireRowAreas != null) Marshal.ReleaseComObject(entireRowAreas);
                if (entireRow != null) Marshal.ReleaseComObject(entireRow);
                if (areas != null) Marshal.ReleaseComObject(areas);
                foreach (var r in rangesFromMap) Marshal.ReleaseComObject(r);
                if (activeSheet != null) Marshal.ReleaseComObject(activeSheet);
            }
            System.Diagnostics.Debug.WriteLine($"DoApplyHighlightingToSelection: END.");
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
                if (initialCount == 0)
                {
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
            Excel.Sheets wss = null;
            Excel.Worksheet ws = null;
            try
            {
                if (this.Application != null)
                {
                    wbs = this.Application.Workbooks;
                    if (wbs != null && wbs.Count > 0)
                    {
                        for (int i = 1; i <= wbs.Count; i++) // COM collections are 1-indexed
                        {
                            wb = wbs[i];
                            wss = wb.Worksheets;
                            if (wss != null)
                            {
                                foreach (object sheetObj in wss)
                                {
                                    if (sheetObj is Excel.Worksheet)
                                    {
                                        ws = (Excel.Worksheet)sheetObj;
                                        RemoveAddinConditionalFormatting(ws);
                                        if (ws != null) { Marshal.ReleaseComObject(ws); ws = null; }
                                    }
                                    else
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
            finally
            {
                // Ensure all COM objects in this scope are released
                if (ws != null) Marshal.ReleaseComObject(ws);
                if (wss != null) Marshal.ReleaseComObject(wss);
                if (wb != null) Marshal.ReleaseComObject(wb);
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