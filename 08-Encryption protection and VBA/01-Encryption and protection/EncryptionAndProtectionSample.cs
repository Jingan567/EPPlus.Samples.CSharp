﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB           Initial release EPPlus 5
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using OfficeOpenXml;
using System.Drawing;
using OfficeOpenXml.Style;

namespace EPPlusSamples.EncryptionProtectionAndVba
{
    public static class EncryptionAndProtection
    {
        public static void Run()
        {
            Console.WriteLine("Running sample 8.1");

            //Create a Sample 19 directory...
            var outputDir = FileUtil.GetDirectoryInfo("8.1-EncryptionAndProtection");

            //create the three FileInfo objects...
            FileInfo templateFile = FileUtil.GetFileInfo(outputDir, "template.xlsx");
            FileInfo answerFile = FileUtil.GetFileInfo(outputDir, "answers.xlsx");
            FileInfo JKAnswerFile = FileUtil.GetFileInfo(outputDir, "JKAnswers.xlsx");

            //Create the template...
            using (ExcelPackage package = new ExcelPackage(templateFile))
            {
                //Lock the workbook totally
                var workbook = package.Workbook;
                workbook.Protection.LockWindows = true;
                workbook.Protection.LockStructure = true;
                workbook.View.SetWindowSize(150, 525, 14500, 6000);
                workbook.View.ShowHorizontalScrollBar = false;
                workbook.View.ShowVerticalScrollBar = false;
                workbook.View.ShowSheetTabs = false;

                //Set a password for the workbookprotection
                workbook.Protection.SetPassword("EPPlus");

                //Encrypt with no password
                package.Encryption.IsEncrypted = true;

                var sheet = package.Workbook.Worksheets.Add("Quiz");
                sheet.View.ShowGridLines = false;
                sheet.View.ShowHeaders = false;
                using(var range=sheet.Cells["A:XFD"])
                {
                    range.Style.Fill.PatternType=ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                    range.Style.Font.Name = "Broadway";
                    range.Style.Hidden = true;
                }

                sheet.Cells["A1"].Value = "Quiz-Sweden";
                sheet.Cells["A1"].Style.Font.Size = 18;
                sheet.Cells["A3"].Value = "Enter your name:";

                sheet.Columns[1].Width = 30;
                sheet.Columns[2].Width = 80;
                sheet.Columns[3].Width = 20;

                sheet.Cells["A7"].Value = "What is the name of the capital of Sweden?";
                sheet.Cells["A9"].Value = "At which place did the Swedish team end up in the Soccer Worldcup 1994?";
                sheet.Cells["A11"].Value = "What is the first name of the famous Swedish inventor/scientist that founded the Nobel-prize?";

                using (var r = sheet.Cells["B3,C7,C9,C11"])
                {
                    r.Style.Fill.BackgroundColor.SetColor(Color.WhiteSmoke);
                    r.Style.Border.Top.Style = ExcelBorderStyle.Dotted;
                    r.Style.Border.Top.Color.SetColor(Color.Black);
                    r.Style.Border.Right.Style = ExcelBorderStyle.Dotted;
                    r.Style.Border.Right.Color.SetColor(Color.Black);
                    r.Style.Border.Bottom.Style = ExcelBorderStyle.Dotted;
                    r.Style.Border.Bottom.Color.SetColor(Color.Black);
                    r.Style.Border.Left.Style = ExcelBorderStyle.Dotted;
                    r.Style.Border.Left.Color.SetColor(Color.Black);
                    r.Style.Locked = false;
                    r.Style.Hidden = false;
                    r.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                }
                sheet.Select("B3");
                sheet.Protection.SetPassword("EPPlus");
                sheet.Protection.AllowSelectLockedCells = false;

                //Options question 1
                var list1 = sheet.Cells["C7"].DataValidation.AddListDataValidation();
                list1.Formula.Values.Add("Bern");
                list1.Formula.Values.Add("Stockholm");
                list1.Formula.Values.Add("Oslo");
                list1.ShowErrorMessage = true;
                list1.Error = "Please select a value from the list";

                var list2 = sheet.Cells["C9"].DataValidation.AddListDataValidation();
                list2.Formula.Values.Add("First");
                list2.Formula.Values.Add("Second");
                list2.Formula.Values.Add("Third");
                list2.ShowErrorMessage = true;
                list2.Error = "Please select a value from the list";

                var list3 = sheet.Cells["C11"].DataValidation.AddListDataValidation();
                list3.Formula.Values.Add("Carl Gustaf");
                list3.Formula.Values.Add("Ingmar");
                list3.Formula.Values.Add("Alfred");
                list3.ShowErrorMessage = true;
                list3.Error = "Please select a value from the list";

                //Save, and the template is ready for use
                package.Save();

                //Quiz-template is done, now create the answer template and encrypt it...
                using (var packageAnswers = new ExcelPackage(package.Stream))       //We use the stream from the template here to get a copy of it.
                {
                    var sheetAnswers = packageAnswers.Workbook.Worksheets[0];
                    sheetAnswers.Cells["C7"].Value = "Stockholm";
                    sheetAnswers.Cells["C9"].Value = "Third";
                    sheetAnswers.Cells["C11"].Value = "Alfred";

                    packageAnswers.Encryption.Algorithm = EncryptionAlgorithm.AES192;   //For the answers we want a little bit stronger encryption
                    packageAnswers.SaveAs(answerFile, "EPPlus");                        //Save and set the password to EPPlus. The password can also be set using packageAnswers.Encryption.Password property
                }

                //Ok, Since this is     qan example we create one user answer...
                using (var packageAnswers = new ExcelPackage(package.Stream))
                {
                    var sheetUser = packageAnswers.Workbook.Worksheets[0];
                    sheetUser.Cells["B3"].Value = "Jan Källman";
                    sheetUser.Cells["C7"].Value = "Bern";
                    sheetUser.Cells["C9"].Value = "Third";
                    sheetUser.Cells["C11"].Value = "Alfred";

                    packageAnswers.SaveAs(JKAnswerFile, "JK");  //We use default encryption here (AES128) and Password JK
                }
            }


            //Now lets correct the user form...
            var packAnswers = new ExcelPackage(answerFile, "EPPlus");    //Supply the password, so the file can be decrypted
            var packUser =  new ExcelPackage(JKAnswerFile, "JK");        //Supply the password, so the file can be decrypted

            var wsAnswers = packAnswers.Workbook.Worksheets[0];
            var wsUser = packUser.Workbook.Worksheets[0];

            //Enumerate the three answers
            foreach (var cell in wsAnswers.Cells["C7,C9,C11"])
            {
                wsUser.Cells[cell.Address].Style.Fill.PatternType = ExcelFillStyle.Solid;
                if (cell.Value.ToString().Equals(wsUser.Cells[cell.Address].Value.ToString(), StringComparison.OrdinalIgnoreCase)) //Correct Answer?
                {
                    wsUser.Cells[cell.Address].Style.Fill.BackgroundColor.SetColor(Color.Green);
                }
                else
                {
                    wsUser.Cells[cell.Address].Style.Fill.BackgroundColor.SetColor(Color.Red);
                }
            }
            packUser.Save();

            packUser.Dispose();
            packAnswers.Dispose();
            packUser.Dispose();

            Console.WriteLine("Sample 8.1 created: {0}", FileUtil.OutputDir.FullName);
            Console.WriteLine();
        }
    }
}
