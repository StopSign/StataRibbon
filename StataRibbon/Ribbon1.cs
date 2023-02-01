using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace StataRibbon
{
    public partial class Ribbon1
    {
        private string fileLoc = @"C:\WBG\runDoLines";
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            label1.Label = "";

        }

        private void editConfiguration_Click(object sender, RibbonControlEventArgs e)
        {
            Process.Start("notepad.exe", fileLoc + @"\rundo51\rundo.ini");
        }

        private void runStataLinesButton_Click(object sender, RibbonControlEventArgs e)
        {
            Range selection = Globals.ThisAddIn.Application.Selection as Range;

            int rows = selection.Rows.Count;
            int columns = selection.Columns.Count;
            if(columns > 1)
            {
                label1.Label = "Error: Multiple columns selected";
                return;
            }
            label1.Label = "";

            string output = "";
            int emptyCellCounter = 0;
            for (int rowIndex = 1; rowIndex <= rows; ++rowIndex)
            {
                Range cell = selection.get_Item(rowIndex, 1) as Range;

                if(cell.Value2 != null)
                {
                    output += cell.Value2+"\n";
                    emptyCellCounter = 0;
                } else
                {
                    output += "\n";
                    emptyCellCounter++;
                    if(emptyCellCounter > 4) //stop running after finding 5 empty rows
                    {
                        break;
                    }
                }
            }

            runStataCode(output);
        }

        private void browseButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                editBox1.Text = folderBrowserDialog1.SelectedPath;
            }
        }


        private void runStata_Click(object sender, RibbonControlEventArgs e)
        {
            var sheet = Globals.ThisAddIn.Application.Worksheets.Cast<Worksheet>()
                                   .SingleOrDefault(w => w.Name == "Code");

            if(sheet == null)
            {
                label1.Label = "Error: No sheet named 'Code'";
                return;
            }

            string output = "";
            int emptyCellCounter = 0;
            int rowCount = 0;
            for(int i = 2; i < sheet.Rows.Count; i++)
            {
                rowCount++;
                string value = (string)(sheet.Cells[i, 3] as Range).Value;
                output += value + "\n";
                if(String.IsNullOrEmpty(value))
                {
                    emptyCellCounter++;
                }
                else
                {
                    emptyCellCounter = 0;
                }
                if(emptyCellCounter > 4)
                {
                    break;
                }
            }
            label1.Label = "Processed " + rowCount + " rows";

            runStataCode(output);
        }

        private void runStataCode(string text)
        {
            File.WriteAllText(fileLoc + @"\temp\temp.txt", text);

            // Use ProcessStartInfo class
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.CreateNoWindow = false;
            startInfo.UseShellExecute = false;
            startInfo.FileName = fileLoc + @"\rundo51\rundo.exe";
            startInfo.WindowStyle = ProcessWindowStyle.Hidden;
            startInfo.Arguments = "\""+ fileLoc + "\\temp\\temp.txt\"";

            try
            {
                // Start the process with the info we specified.
                // Call WaitForExit and then the using statement will close.
                using(Process exeProcess = Process.Start(startInfo))
                {
                    exeProcess.WaitForExit();
                }
            }
            catch(Exception e)
            {
                // Log error.
                label1.Label = "Error: " + e.Message;
            }
        }
        private void makeDoFile_Click(object sender, RibbonControlEventArgs e)
        {
            string folderPath = editBox1.Text;
            string fileName = editBox2.Text;

            if(folderPath == null || (folderPath is string && String.IsNullOrEmpty(folderPath)))
            {
                label1.Label = "Error: Folder Path is blank!";
                return;
            }
            if(fileName == null || (fileName is string && String.IsNullOrEmpty(fileName)))
            {
                label1.Label = "Error: File Name is blank!";
                return;
            }

            var headerSheet = Globals.ThisAddIn.Application.Worksheets.Cast<Worksheet>().SingleOrDefault(w => w.Name == "Header");
            if(headerSheet == null)
            {
                label1.Label = "Error: No sheet named 'Header'";
                return;
            }
            var codeSheet = Globals.ThisAddIn.Application.Worksheets.Cast<Worksheet>().SingleOrDefault(w => w.Name == "Code");
            if(codeSheet == null)
            {
                label1.Label = "Error: No sheet named 'Code'";
                return;
            }

            string output = "/*****************************************************************************************************\n******************************************************************************************************\n**                                                                                                  **\n**                       INTERNATIONAL INCOME DISTRIBUTION DATABASE (I2D2)                          **\n**                                                                                                  **\n";
            int emptyCellCounter = 0;
            for(int i = 1; i < headerSheet.Rows.Count; i++)
            {
                var title = (headerSheet.Cells[i, 1] as Range).Value;
                var info = (headerSheet.Cells[i, 2] as Range).Value;
                var tag = (headerSheet.Cells[i, 3] as Range).Value;
                if((title == null || (title is string && String.IsNullOrEmpty(title))) &&
                    (tag == null || (tag is string && String.IsNullOrEmpty(tag))))
                {
                    emptyCellCounter++;
                    if(emptyCellCounter > 4)
                    {
                        break;
                    }
                }
                else
                {
                    output += "*" + title + "\t\t\t" + info + "\t\t\t"+tag+"\n";
                    emptyCellCounter = 0;
                }
            }
            output += "**                                                                                                  **\n******************************************************************************************************\n*****************************************************************************************************/\n";

            emptyCellCounter = 0;
            int rowCount = 0;
            for(int i = 2; i < codeSheet.Rows.Count; i++)
            {
                rowCount++;
                var block = (codeSheet.Cells[i, 1] as Range).Value;
                var variableLongName = (codeSheet.Cells[i, 2] as Range).Value;
                var stataCode = (codeSheet.Cells[i, 3] as Range).Value;
                var altCode = (codeSheet.Cells[i, 4] as Range).Value;
                var largeComment = (codeSheet.Cells[i, 5] as Range).Value;
                var codeExplanation = (codeSheet.Cells[i, 6] as Range).Value;
                var originalQuestion = (codeSheet.Cells[i, 7] as Range).Value;
                var document = (codeSheet.Cells[i, 8] as Range).Value;


                if(stataCode == null || (stataCode is string && String.IsNullOrEmpty(stataCode)))
                {
                    emptyCellCounter++;
                    output += "\n";
                    if(emptyCellCounter > 4)
                    {
                        break;
                    }
                }
                else
                {
                    emptyCellCounter = 0;
                }

                if(block is string && !String.IsNullOrEmpty(block))
                {
                    output += "\n/*****************************************************************************************************\n*                                                                                                    *\n                                   ";
                    output += block;
                    output += "\n*                                                                                                    *\n*****************************************************************************************************/\n\n\n";
                }
                if(variableLongName is string && !String.IsNullOrEmpty(variableLongName))
                {
                    output += "\n** " + variableLongName + "\n";
                }
                if(largeComment is string && !String.IsNullOrEmpty(largeComment))
                {
                    output += "\n/*\n" + largeComment + "\n*/\n";
                }
                if(originalQuestion is string && !String.IsNullOrEmpty(originalQuestion))
                {
                    output += "\n/* Original Question\n" + originalQuestion + "\n*/\n";
                }
                if(stataCode is string && !String.IsNullOrEmpty(stataCode))
                {
                    output += "\t" + stataCode;
                    if(codeExplanation is string && !String.IsNullOrEmpty(codeExplanation))
                    {
                        output += " // " + codeExplanation;
                    }
                    output += "\n";
                }
                else
                {
                    if(codeExplanation is string && !String.IsNullOrEmpty(codeExplanation))
                    {
                        output += " // " + codeExplanation + "\n";
                    }
                }
                if(altCode is string && !String.IsNullOrEmpty(altCode))
                {
                    output += "* " + altCode + "\n";
                }
                if(document is string && !String.IsNullOrEmpty(document))
                {
                    output += "\n* Document: " + document + "\n";
                }
            }
            output += "******************************  END OF DO-FILE  *****************************************************/";

            label1.Label = "Processed " + rowCount + " rows";

            if(!fileName.EndsWith(".do"))
            {
                fileName += ".do";
            }
            File.WriteAllText(folderPath + "\\" + fileName, output);
        }
    }
}
