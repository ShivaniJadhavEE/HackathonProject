using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace HackathonProject
{
    public partial class RibbonData
    {

        private Microsoft.Office.Interop.Excel.ListObject dataTable;
        private Microsoft.Office.Interop.Excel.ListObject flagTable;
        public Microsoft.Office.Interop.Excel.ListObject tableAfterChange;
        public List<List<object>> valuesAfterChange;
        public List<List<object>> flagValuesAfter;
       
        private void RibbonData_Load(object sender, RibbonUIEventArgs e)
        {
         
            
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Microsoft.Office.Interop.Excel.Worksheet worksheet = Globals.ThisAddIn.ActiveWorksheet;

            ListObjects tables = worksheet.ListObjects;
            bool tableExists = false;

            // Now, you can iterate through the tables in the collection
            foreach (Microsoft.Office.Interop.Excel.ListObject table in tables)
            {
                // Access information about each table, such as its name or data
                if (table.Name.Equals("datatable"))
                {
                    tableAfterChange = table;
                    tableExists = true;
                }
                // Do something with the table data or properties
                if (table.Name.Equals("flagtable"))
                {
                    flagTable = table;
                }
            }

            //get data after change
            if (tableExists == true)
            {
                valuesAfterChange = new List<List<object>>();
                foreach (Microsoft.Office.Interop.Excel.ListRow row in tableAfterChange.ListRows)
                {
                    List<object> rowData = new List<object>();
                    foreach (Microsoft.Office.Interop.Excel.ListColumn column in tableAfterChange.ListColumns)
                    {
                        if (row.Index == 1)
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
                            if (row.Range.Cells[1, column.Index].Value2 is string)
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
                    valuesAfterChange.Add(rowData);
                }
            }

            //get data for flag columns          
                flagValuesAfter = new List<List<object>>();
                foreach (Microsoft.Office.Interop.Excel.ListRow row in flagTable.ListRows)
                {
                    List<object> rowData = new List<object>();
                    foreach (Microsoft.Office.Interop.Excel.ListColumn column in flagTable.ListColumns)
                    {                      
                                rowData.Add(row.Range.Cells[1, column.Index].Value2);                     
                    }
                flagValuesAfter.Add(rowData);
                }

                if(AreEqual(Globals.ThisAddIn.valuesBeforeChange, valuesAfterChange)&&AreEqual(flagValuesAfter, Globals.ThisAddIn.flagValuesBefore))
                {
                    MessageBox.Show("Values of data table and flag table are matching.", "Matching Values", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            else
            {
                GenerateCsv(flagValuesAfter, valuesAfterChange);
            }
            
        }

        private bool AreEqual(List<List<object>> list1, List<List<object>> list2)
        {
            if (list1.Count != list2.Count)
                return false;

            for (int i = 0; i < list1.Count; i++)
            {
                if (!list1[i].SequenceEqual(list2[i]))
                    return false;
            }

            return true;
        }

        public void GenerateCsv(List<List<object>> flagValuesAfter, List<List<object>> valuesAfterChange)
        {
            List<object> dateTimeData = new List<object>();
            List<object> objectListfxd = new List<object>();
            List<object> objectListcommit=new List<object>();
            List<List<object>> commitData= new List<List<object>>();
            List<List<object>> fixLoadData = new List<List<object>>();
            
            for (int i=0;i<valuesAfterChange.Count;i++)
            {
                for(int j = 0; j < valuesAfterChange[i].Count;j++)
                {
                    if (i == 0 && valuesAfterChange[i][j]!=null)
                    {
                        dateTimeData.Add(valuesAfterChange[0][j]);
                    }
                    /*if(j==0&& valuesAfterChange[i][j] != null)
                    {
                        objectList.Add(valuesAfterChange[i][0]);
                    }*/
                }
            }
            objectListfxd.Insert(0, "DateTime");
            objectListcommit.Insert(0, "DateTime");
            string flag = "";
            for (int i = 1; i < flagValuesAfter.Count; i++)
            {
                if (flagValuesAfter[i] != null)
                {
                    if (flagValuesAfter[i][2].Equals("Yes"))
                    {
                        flag = "FXD";
                        objectListfxd.Add(valuesAfterChange[i][0]);
                        fixLoadData = GetFinalData(valuesAfterChange, fixLoadData, flag,i);
                    }
                    else
                    {
                        if(flagValuesAfter[i][1].Equals("Yes"))
                        {
                            flag = "MRN";
                            objectListcommit.Add(valuesAfterChange[i][0]);
                            commitData = GetFinalData(valuesAfterChange, commitData, flag,i);
                        }
                        else
                        {
                            flag = "ECO";
                            commitData= GetFinalData(valuesAfterChange, commitData, flag, i);
                            fixLoadData =GetFinalData(valuesAfterChange, fixLoadData, flag,i);
                            objectListfxd.Add(valuesAfterChange[i][0]);
                            objectListcommit.Add(valuesAfterChange[i][0]);

                        }
                    }

                    
                   
                }
            }
            string filePath1 = "C:\\CSVDataHack\\commit.csv";
            string filePath2 = "C:\\CSVDataHack\\fix.csv";
            if (fixLoadData.Count() > 0)
            {
                fixLoadData = TransposeData(fixLoadData);
                DownloadCSV(fixLoadData, objectListfxd, dateTimeData,filePath2);


            }
            if (commitData.Count() > 0)
            {
                commitData = TransposeData(commitData);
                DownloadCSV(commitData, objectListcommit, dateTimeData,filePath1);
            }

        }

        public List<List<object>> GetFinalData(List<List<object>> valuesAfterChange, List<List<object>> ToList, string flag,int index)
        {
            List<object>temp= new List<object>();
          
            if (flag!=null&& !flag.Equals("ECO")) {

                for (int i = 1; i < valuesAfterChange.Count; i++)
                {
                    for (int j = 2; j < valuesAfterChange[i].Count; j++)
                    {
                        if(i==index)
                        {
                            object a = valuesAfterChange[i][j];
                            object b = Globals.ThisAddIn.valuesBeforeChange[i][j];
                            if (valuesAfterChange[i][j].Equals(Globals.ThisAddIn.valuesBeforeChange[i][j]))
                            {

                                valuesAfterChange[i][j] = -1;
                            }
                            temp.Add(valuesAfterChange[i][j]);
                        }
                       
                    }
                }
                ToList.Add(temp);
            }
            else
            {
                for (int i = 1; i < valuesAfterChange.Count; i++)
                {
                    for (int j = 2; j < valuesAfterChange[i].Count; j++)
                    {
                        if(i==index)
                        {
                            valuesAfterChange[i][j] = -1;
                            temp.Add(valuesAfterChange[i][j]);

                        }
                       
                           
                       
                    }
                }
                ToList.Add(temp);
            }
           
            
            return ToList;
        }

        public List<List<object>> TransposeData(List<List<object>> valuesAfterChange)
        {
            
           
            int numRows = valuesAfterChange.Count;
            int numCols = valuesAfterChange[0].Count;

            // Initialize transposeData with the transposed dimensions
            List<List<object>> transposeData = new List<List<object>>();
            for (int j = 0; j < numCols; j++)
            {
                transposeData.Add(new List<object>());
            }

            // Transpose the data
            for (int j = 0; j < numCols; j++)
            {
                for (int i = 0; i < numRows; i++)
                {
                    transposeData[j].Add(valuesAfterChange[i][j]);
                }
            }

            return transposeData;

        }

        public void DownloadCSV(List<List<object>> valuesAfterChange,List<object>objectData,List<object>dateTimeData,string outputPath)
        {
            List<List<object>> finalDataCSV= new List<List<object>>();
            //valuesAfterChange[1].Clear();
            finalDataCSV.Add(objectData);
           for(int i=0;i<valuesAfterChange.Count;i++)
            {
                finalDataCSV.Add(valuesAfterChange[i]);
            }
           for(int i = 1; i < finalDataCSV.Count; i++)
            {
                finalDataCSV[i].Insert(0, dateTimeData[i-1]);
            }
           
            using (StreamWriter sw = new StreamWriter(outputPath))
            {
                // Write data to the CSV file
                foreach (List<object> row in finalDataCSV)
                {
                    // Join the values in the row with commas and write to the file
                    string rowString = string.Join(",", row);
                    sw.WriteLine(rowString);
                }
            }
        }

        /* private void checkBox1_Click(object sender, RibbonControlEventArgs e)
         {
             string customTableName = "hackTable";
             Microsoft.Office.Interop.Excel.Worksheet worksheet = Globals.ThisAddIn.ActiveWorksheet;


             if (((RibbonCheckBox)sender).Checked)
             {

                 // Add a ListObject (Excel Table) to the worksheet
                 customTable = worksheet.ListObjects.Add(XlListObjectSourceType.xlSrcRange, worksheet.Range["G1"], Type.Missing, XlYesNoGuess.xlYes);
                 customTable.Name = customTableName;
                 customTable.ListColumns[1].Name = "ECO";


                 ListColumn column2 = customTable.ListColumns.Add();
                 column2.Name = "MRN";


                 ListColumn column3 = customTable.ListColumns.Add();
                 column3.Name = "FXD";

                 customTable.ListRows.Add();
                 customTable.ListRows[1].Range[1].Value2 = "No";
                 customTable.ListRows[1].Range[2].Value2 = "No";
                 customTable.ListRows[1].Range[3].Value2 = "No";
             }
             else
             {
                 if (customTable != null)
                 {
                     customTable.Delete();
                     customTable = null; // Set the reference to null to indicate the table is removed
                 }
             }



         }*/
    }
}
