using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace HackathonProject
{
    public partial class ThisAddIn
    {
        public Worksheet ActiveWorksheet { get; private set; }
        public Microsoft.Office.Interop.Excel.ListObject tableBeforeChange;
        public Microsoft.Office.Interop.Excel.ListObject flagTableBefore;
        public string mainTableName;
        public List<List<object>> valuesBeforeChange;
        public List<List<object>> flagValuesBefore;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            ActiveWorksheet = (Worksheet)(Globals.ThisAddIn.Application.ActiveSheet as Worksheet);

            // Set up event handlers for sheet activation and deactivation
            Globals.ThisAddIn.Application.SheetActivate += Application_SheetActivate;
            Globals.ThisAddIn.Application.SheetDeactivate += Application_SheetDeactivate;
            Globals.ThisAddIn.Application.WorkbookOpen += Application_WorkbookOpen;
        }

        private void Application_WorkbookOpen(object sh)
        {
            // Call HandleFileSelection when a workbook is opened
            HandleFileSelection();
        }
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }
        private void Application_SheetActivate(object sh)
        {
            // The active sheet has changed; update ActiveWorksheet
            ActiveWorksheet = (sh as Worksheet);
        }

        private void Application_SheetDeactivate(object sh)
        {
            // A sheet has lost focus; you can handle this event if needed
        }

        private void HandleFileSelection()
        {
            // Ensure there is an active workbook
            if (Globals.ThisAddIn.Application.ActiveWorkbook != null)
            {
                // Update ActiveWorksheet with the initially active sheet
                ActiveWorksheet = Globals.ThisAddIn.Application.ActiveSheet as Worksheet;

                ListObjects tables = ActiveWorksheet.ListObjects;
                bool tableExists = false;

                // Now, you can iterate through the tables in the collection
                foreach (Microsoft.Office.Interop.Excel.ListObject table in tables)
                {
                    // Access information about each table, such as its name or data
                    string tableName = table.Name;
                    if (table.Name.Equals("datatable"))
                    {
                        tableBeforeChange = table;
                        tableExists = true;
                    }
                    // Do something with the table data or properties
                    if (table.Name.Equals("flagtable"))
                    {
                        flagTableBefore = table;
                    }



                    // Do something with the table data or properties
                }
                if(tableExists==true)
                {
                    valuesBeforeChange = new List<List<object>>();
                    foreach (Excel.ListRow row in tableBeforeChange.ListRows)
                    {
                        List<object> rowData = new List<object>();
                        foreach (Excel.ListColumn column in tableBeforeChange.ListColumns)
                        {
                            if(row.Index==1)
                            {
                                if (row.Range.Cells[1, column.Index].Value2 != null)
                                {
                                    double excelDateValue = row.Range.Cells[1, column.Index].Value2; // Example: 44253.0
                                    DateTime dateTimeValue = DateTime.FromOADate(excelDateValue);
                                    string formattedDateTimeString = dateTimeValue.ToString("dd-MM-yyyy HH:mm");
                                    rowData.Add(formattedDateTimeString);
                                }
                                else
                                {
                                    rowData.Add(row.Range.Cells[1, column.Index].Value2);
                                }
                            }
                            else
                            {
                                if( row.Range.Cells[1, column.Index].Value2 is string)
                                {
                                    rowData.Add(row.Range.Cells[1, column.Index].Value2);
                                }
                                else
                                {
                                    double inputValue = row.Range.Cells[1, column.Index].Value2;
                                    float formattedFloatValue = (float)Math.Round(inputValue, 1);
                                    rowData.Add(formattedFloatValue);
                                }
                                
                            }
                           
                        }
                        valuesBeforeChange.Add(rowData);
                    }
                }
                flagValuesBefore = new List<List<object>>();
                foreach (Microsoft.Office.Interop.Excel.ListRow row in flagTableBefore.ListRows)
                {
                    List<object> rowData = new List<object>();
                    foreach (Microsoft.Office.Interop.Excel.ListColumn column in flagTableBefore.ListColumns)
                    {
                        rowData.Add(row.Range.Cells[1, column.Index].Value2);
                    }
                    flagValuesBefore.Add(rowData);
                }
            }
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
