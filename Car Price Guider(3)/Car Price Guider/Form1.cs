﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using System.Threading;
using System.Net;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using FolderSelect;
using Microsoft.VisualBasic.FileIO;
using System.Reflection;

namespace Car_Price_Guider
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            
            LinkLabel.Link link = new LinkLabel.Link();
            link.LinkData = "http://www.honestjohn.co.uk/used-prices/";
            linkLabel1.Links.Add(link);

            txtBox_SamplePageDir.Text = @"H:\Visual Studio 2010 Projects\Car Price Guider(2)\Car Price Guider\Sample Pages\2013-09-28T21-09-22";
            txtBox_SamplePageDir.Text = @"D:\Users\David\Documents\Stuff To Keep Synced\Visual Studio 2010\Projects\Car Price Guider(2)\Sample Pages\2013-09-28T21-09-22";
            txtBox_SamplePageDir.Text = @"D:\Users\David\Documents\Stuff To Keep Synced\Visual Studio 2010\Projects\Car Price Guider(2)\Sample Pages\";
            
            //txtBox_SamplePageDir.Text = @"D:\Users\David\Documents\Stuff To Keep Synced\Visual Studio 2010\Projects\Car Price Guider(2)\Sample Pages\";

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string baseURL = @"http://www.honestjohn.co.uk/used-prices/";
            string args = "?q=";
            string detailsToFind = "Ford Mondeo 2004 Automatic Estate Ghia X TDCi";
            detailsToFind = "renault Clio 2003 extreme";

            foreach (string currLine in textBox5.Lines)
            {
                
                detailsToFind = currLine.Trim();

                if (detailsToFind != "")
                {
                    string fullURL = baseURL + args + detailsToFind;

                    if (false)
                    {
                        Console.WriteLine(detailsToFind);

                        ProcessStartInfo startInfo = new ProcessStartInfo("IExplore.exe");

                        startInfo.Arguments = fullURL;

                        Process.Start(startInfo);
                    }

                    if (true)
                    {
                        Console.WriteLine(fullURL);
                        ReadWebPagePrices(fullURL, detailsToFind);
                    }

                    int secondsToWait = 30;

                    Thread.Sleep(secondsToWait * 1000);
                }
                
            } // end foreach

        }

        private void SaveWebPage(string url, string fname)
        {
            //string url = @"D:\Users\David\Documents\Visual Studio 2010\Projects\Car Price Guider\Sample Pages\Ford Mondeo 2004 Automatic Estate Ghia X TDCi Price Guide   Honest John.htm";
            //url = @"D:\Users\David\Documents\Visual Studio 2010\Projects\Car Price Guider\Sample Pages\VAUXHALL ASTRA MERIT AUTO - 1389cc 5dr Hatchback 1995 Petrol Auto Price Guide   Honest John.htm";

            //List<string> src;

            var req = (HttpWebRequest)WebRequest.Create(url);
            //var req = (FileWebRequest)WebRequest.Create(url);
            req.Method = "GET";

            using (WebResponse odpoved = req.GetResponse())
            {
                var htmlDoc = new HtmlAgilityPack.HtmlDocument();
                //htmlDoc.Load(odpoved.GetResponseStream());


                Console.WriteLine(DateTime.Now.ToShortDateString());
                var fileStream = File.Create(fname);

                Stream str = odpoved.GetResponseStream();
                str.CopyTo(fileStream);
                fileStream.Close();
            }

            return;
        }

        private void ReadWebPagePrices(string url, string fname)
        {
            //string url = @"D:\Users\David\Documents\Visual Studio 2010\Projects\Car Price Guider\Sample Pages\Ford Mondeo 2004 Automatic Estate Ghia X TDCi Price Guide   Honest John.htm";
            //url = @"D:\Users\David\Documents\Visual Studio 2010\Projects\Car Price Guider\Sample Pages\VAUXHALL ASTRA MERIT AUTO - 1389cc 5dr Hatchback 1995 Petrol Auto Price Guide   Honest John.htm";
            
            //List<string> src;

            var req = (HttpWebRequest)WebRequest.Create(url);
            //var req = (FileWebRequest)WebRequest.Create(url);
            req.Method = "GET";

            using (WebResponse odpoved = req.GetResponse())
            {
                var htmlDoc = new HtmlAgilityPack.HtmlDocument();
                //htmlDoc.Load(odpoved.GetResponseStream());


                Console.WriteLine(DateTime.Now.ToShortDateString());
                var fileStream = File.Create("D:\\" + fname + ".html");

                Stream str = odpoved.GetResponseStream();
                str.CopyTo(fileStream);
                fileStream.Close();
            }

            return;

            var req1 = (FileWebRequest)WebRequest.Create("D:\\test3.html");
            req1.Method = "GET";

            using (WebResponse odpoved = req1.GetResponse())
            {
                var htmlDoc = new HtmlAgilityPack.HtmlDocument();
                htmlDoc.Load(odpoved.GetResponseStream());
                //var fileStream = File.Create("C:\\test.html");
                //myOtherObject.InputStream.Seek(0, SeekOrigin.Begin);
                //myOtherObject.InputStream.CopyTo(fileStream);
                //fileStream.Close();

                //Console.WriteLine(htmlDoc.te);



                string minPrice = GetPathText(htmlDoc, "//*[@id=\"prices\"]/div[3]/div/div/div[3]/p[1]/b[1]");
                string maxPrice = GetPathText(htmlDoc, "//*[@id=\"prices\"]/div[3]/div/div/div[3]/p[1]/b[2]");

                string dealerMinPrice = GetPathText(htmlDoc, "//*[@id=\"prices\"]/div[3]/div/div/table/tbody/tr[1]/td[3]/b[1]");
                //*[@id=\"prices\"]/div[3]/div/div/table/tbody/tr[1]/td[3]/b[1]
                //*[@id=\"prices\"]/div[3]/div/div/table/tbody/tr[1]/td[3]/b[1]

                string dealerMaxPrice = GetPathText(htmlDoc, "//*[@id=\"prices\"]/div[3]/div/div/table/tbody/tr[1]/td[3]/b[2]");


                string privateSellerMinPrice = GetPathText(htmlDoc, "//*[@id=\"prices\"]/div[3]/div/div/table/tbody/tr[3]/td[3]/b[1]");
                //string privateSellerMaxPrice = GetPathText(htmlDoc, "//*[@id=\"prices\"]/div[3]/div/div/table/tbody/tr[3]/td[3]/b[2]");
                //string privateSellerMinPrice = GetPathText(htmlDoc, "//*[@id=\"prices\"]/div[0]/");
                string privateSellerMaxPrice = GetPathText(htmlDoc, "//*[@id=\"prices\"]/div[3]/div/div/table/tbody/tr[3]/td[3]/b[2]");


                string partExPrice = GetPathText(htmlDoc, "//*[@id=\"prices\"]/div[3]/div/div/table/tbody/tr[5]/td[3]/b");
                //string partExMaxPrice = GetPathText(htmlDoc, "//*[@id=\"prices\"]/div[3]/div/div/table/tbody/tr[1]/td[3]/b[1]");

                Console.WriteLine("Most adverts are between " + minPrice + " and " + maxPrice);

                Console.WriteLine("Dealer Prices are between " + dealerMinPrice + " and " + dealerMaxPrice);

                Console.WriteLine("Private Prices are between " + privateSellerMinPrice + " and " + privateSellerMaxPrice);

                Console.WriteLine("Part ex price is " + partExPrice);

            } // end using
        }

        private void button2_Click(object sender, EventArgs e)
        {

            string url = @"D:\Users\David\Documents\Visual Studio 2010\Projects\Car Price Guider\Sample Pages\Ford Mondeo 2004 Automatic Estate Ghia X TDCi Price Guide   Honest John.htm";

            List<string> src;

            //var req = (HttpWebRequest)WebRequest.Create(url);
            var req = (FileWebRequest)WebRequest.Create(url);
            req.Method = "GET";

            using (WebResponse odpoved = req.GetResponse())
            {
                var htmlDoc = new HtmlAgilityPack.HtmlDocument();
                htmlDoc.Load(odpoved.GetResponseStream());

                //var nodes = htmlDoc.DocumentNode.SelectNodes("//*[@id=\"prices\"]/div[3]/div/div/div[3]/p[1]/b[1]");
                //src = new List<string>(nodes.Count);

                //if (nodes != null)
                //{
                //    foreach (var node in nodes)
                //    {
                //        if (node.Id != null)
                //        {
                //            src.Add(node.Id);
                //            Console.WriteLine("Most adverts are between " + node.InnerText);
                //        }
                        
                //    }
                //}

                string minPrice = GetPathText(htmlDoc, "//*[@id=\"prices\"]/div[3]/div/div/div[3]/p[1]/b[1]");
                string maxPrice = GetPathText(htmlDoc, "//*[@id=\"prices\"]/div[3]/div/div/div[3]/p[1]/b[2]");

                string dealerMinPrice = GetPathText(htmlDoc, "//*[@id=\"prices\"]/div[3]/div/div/table/tbody/tr[1]/td[3]/b[1]");
                string dealerMaxPrice = GetPathText(htmlDoc, "//*[@id=\"prices\"]/div[3]/div/div/table/tbody/tr[1]/td[3]/b[2]");

                string privateSellerMinPrice = GetPathText(htmlDoc, "//*[@id=\"prices\"]/div[3]/div/div/table/tbody/tr[3]/td[3]/b[1]");
                string privateSellerMaxPrice = GetPathText(htmlDoc, "//*[@id=\"prices\"]/div[3]/div/div/table/tbody/tr[3]/td[3]/b[2]");

                string partExPrice = GetPathText(htmlDoc, "//*[@id=\"prices\"]/div[3]/div/div/table/tbody/tr[5]/td[3]/b");
                //string partExMaxPrice = GetPathText(htmlDoc, "//*[@id=\"prices\"]/div[3]/div/div/table/tbody/tr[1]/td[3]/b[1]");

                Console.WriteLine("Most adverts are between " + minPrice + " and " + maxPrice);

                Console.WriteLine("Dealer Prices are between " + dealerMinPrice + " and " + dealerMaxPrice);

                Console.WriteLine("Private Prices are between " + privateSellerMinPrice + " and " + privateSellerMaxPrice);

                Console.WriteLine("Part ex price is " + partExPrice);

            } // end using



        }

        private void button3_Click(object sender, EventArgs e)
        {

            var req1 = (FileWebRequest)WebRequest.Create("D:\\TOYOTA AVENSIS COLOUR CTION VVTI - 1794cc 5dr Hatchback 2006 Petrol Manual.html");
            req1.Method = "GET";

            using (WebResponse odpoved = req1.GetResponse())
            {
                var htmlDoc = new HtmlAgilityPack.HtmlDocument();
                htmlDoc.Load(odpoved.GetResponseStream());
                //var fileStream = File.Create("C:\\test.html");
                //myOtherObject.InputStream.Seek(0, SeekOrigin.Begin);
                //myOtherObject.InputStream.CopyTo(fileStream);
                //fileStream.Close();

                //Console.WriteLine(htmlDoc.te);



                string minPrice = GetPathText(htmlDoc, "//*[@id=\"prices\"]/div[3]/div/div/div[3]/p[1]/b[1]");
                string maxPrice = GetPathText(htmlDoc, "//*[@id=\"prices\"]/div[3]/div/div/div[3]/p[1]/b[2]");

                string dealerMinPrice = GetPathText(htmlDoc, "//*[@id=\"prices\"]/div[3]/div/div/table/tbody/tr[1]/td[3]/b[1]");
                                                              //*[@id=\"prices\"]/div[3]/div/div/table/tbody/tr[1]/td[3]/b[1]
                //*[@id=\"prices\"]/div[3]/div/div/table/tbody/tr[1]/td[3]/b[1]
                //*[@id=\"prices\"]/div[3]/div/div/table/tbody/tr[1]/td[3]/b[1]

                string dealerMaxPrice = GetPathText(htmlDoc, "//*[@id=\"prices\"]/div[3]/div/div/table/tbody/tr[1]/td[3]/b[2]");


                //string privateSellerMinPrice = GetPathText(htmlDoc, "//*[@id=\"prices\"]/div[3]/div/div/table/tbody/tr[3]/td[3]/b[1]");
                //string privateSellerMaxPrice = GetPathText(htmlDoc, "//*[@id=\"prices\"]/div[3]/div/div/table/tbody/tr[3]/td[3]/b[2]");
                string privateSellerMinPrice = GetPathText(htmlDoc, "//*[@class=\"pricepoints\"]/tr[1]/td[3]/b[1]");
                string privateSellerMaxPrice = GetPathText(htmlDoc, "//*[@id=\"prices\"]/div[3]/div/div/table/tbody/tr[3]/td[3]/b[2]");


                string partExPrice = GetPathText(htmlDoc, "//*[@id=\"prices\"]/div[3]/div/div/table/tbody/tr[5]/td[3]/b");
                //string partExMaxPrice = GetPathText(htmlDoc, "//*[@id=\"prices\"]/div[3]/div/div/table/tbody/tr[1]/td[3]/b[1]");

                Console.WriteLine("Most adverts are between " + minPrice + " and " + maxPrice);

                Console.WriteLine("Dealer Prices are between " + dealerMinPrice + " and " + dealerMaxPrice);

                Console.WriteLine("Private Prices are between " + privateSellerMinPrice + " and " + privateSellerMaxPrice);

                Console.WriteLine("Part ex price is " + partExPrice);

            } // end using

        }

        private CarValuation_HonestJohn ProcessFile(string fname)
        {
            //string fname = "";
            //fname = "RENAULT CLIO DYNAMIQUE 16V - 1149cc 3dr Hatchback 2004 Petrol Manual.html";
            //fname = "TOYOTA AVENSIS COLOUR CTION VVTI - 1794cc 5dr Hatchback 2006 Petrol Manual.html";

            CarValuation_HonestJohn returnPrices = new CarValuation_HonestJohn();

            var req1 = (FileWebRequest)WebRequest.Create(fname);
            req1.Method = "GET";

            using (WebResponse odpoved = req1.GetResponse())
            {
                var htmlDoc = new HtmlAgilityPack.HtmlDocument();
                htmlDoc.Load(odpoved.GetResponseStream());



                /*
                 
                 * Below are working
                //*[@class="pricePoints"]/tr[1]/td[1]/b[1] - dealer
                //*[@class="pricePoints"]/tr[1]/td[3]/b[1] - 4050
                //*[@class="pricePoints"]/tr[1]/td[3]/b[2] - 5200
                 * 
                 */

                string minPrice = GetPathText(htmlDoc, "//*[@id=\"prices\"]/div[3]/div/div/div[3]/p[1]/b[1]");
                string maxPrice = GetPathText(htmlDoc, "//*[@id=\"prices\"]/div[3]/div/div/div[3]/p[1]/b[2]");

                string typeName1 = GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]/tr[1]/td[1]/b[1]");
                string typeName2 = GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]/tr[2]/td[1]/b[1]");
                string typeName3 = GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]/tr[3]/td[1]/b[1]");
                string typeName4 = GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]/tr[4]/td[1]/b[1]");
                string typeName5 = GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]/tr[5]/td[1]/b[1]");


                string dealerMinPrice = GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]/tr[1]/td[3]/b[1]");
                string dealerMaxPrice = GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]/tr[1]/td[3]/b[2]");

                string dealerFranchisedExpectedPrice = GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]/tr[2]/td[3]/b[1]");

                string dealerIndependantExpectedPrice = GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]/tr[3]/td[3]/b[1]");

                string privateSellerMinPrice = GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]/tr[4]/td[3]/b[1]");
                string privateSellerMaxPrice = GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]/tr[4]/td[3]/b[2]");

                string privateSellerExpectedPrice = GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]/tr[5]/td[2]/b[1]");

                string partExPrice = GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]/tr[6]/td[3]/b[1]");

                Console.WriteLine("Most adverts are between " + minPrice + " and " + maxPrice);


                Console.WriteLine("Type1: " + typeName1);
                Console.WriteLine("Type2: " + typeName2);
                Console.WriteLine("Type3: " + typeName3);
                Console.WriteLine("Type4: " + typeName4);
                Console.WriteLine("Type5: " + typeName5);

                /*
                Console.WriteLine("Dealer Prices are between " + dealerMinPrice + " and " + dealerMaxPrice);

                Console.WriteLine("Franchised Dealer expect to pay " + dealerFranchisedExpectedPrice);

                Console.WriteLine("Independant Dealer expect to pay " + dealerIndependantExpectedPrice);

                Console.WriteLine("Private Prices are between " + privateSellerMinPrice + " and " + privateSellerMaxPrice);

                Console.WriteLine("Private seller expect to pay " + privateSellerExpectedPrice);

                Console.WriteLine("Part ex price is " + partExPrice);
                Console.WriteLine();
                */

                returnPrices.DealerMinPrice = dealerMinPrice;
                returnPrices.DealerMaxPrice = dealerMaxPrice;

                returnPrices.DealerFranchisedExpectedPrice = dealerFranchisedExpectedPrice;
                returnPrices.DealerIndependantExpectedPrice = dealerIndependantExpectedPrice;

                returnPrices.PrivateSellerMinPrice = privateSellerMinPrice;
                returnPrices.PrivateSellerMaxPrice = privateSellerMaxPrice;

                returnPrices.PrivateSellerExpectedPrice = privateSellerExpectedPrice;

                returnPrices.PartExPrice = partExPrice;



                // All this below will format the text in 1 block
                string allPriceText = GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]").Trim(); // this works atm

                allPriceText = allPriceText.Trim();
                
                if (allPriceText != "")
                {
                    allPriceText = allPriceText.Replace("\r\n", " ");
                    allPriceText = allPriceText.Replace("\r", " ");
                    allPriceText = allPriceText.Replace("\n", " ");
                    allPriceText = allPriceText.Replace("\t", " ");
                    allPriceText = allPriceText.Replace("At", "\nAt");
                    allPriceText = allPriceText.Replace("From", "\nFrom");
                    allPriceText = allPriceText.Replace("If selling", "\nIf selling");

                    for (int i = 0; i < 200; i++)
                    {
                        allPriceText = allPriceText.Replace("  ", " ");
                    }

                    allPriceText = allPriceText.Replace("Â", " ");
                    allPriceText = allPriceText.Replace("&nbsp;", " ");

                    returnPrices.AllPriceText = allPriceText;
                    
                    //Console.WriteLine(minPrice);
                }
                
            } // end using

            return returnPrices;

        }

        private void button4_Click(object sender, EventArgs e)
        {

            string[] filePaths = Directory.GetFiles(@"D:\AuctionFiles");

            foreach (var item in filePaths)
            {

                Console.WriteLine(item);

                ProcessFile(item);

            }

        }

        private void button5_Click(object sender, EventArgs e)
        {

            DateTime startDateTime = DateTime.Now;

            string workingDir = @"D:\Users\David\Documents\Visual Studio 2010\Projects\Car Price Guider\Sample Pages";

            string workingDirName = "";

            workingDirName = startDateTime.ToString("yyyy-mm-ddTHH-mm-ss");


            workingDir = Path.Combine(workingDir, workingDirName);

            if (Directory.Exists(workingDir))
            {
                MessageBox.Show("Directory exists");
                return;
            }
            else
            {
                Directory.CreateDirectory(workingDir);
            }

            string filename = @"D:\Users\David\Documents\Visual Studio 2010\Projects\Car Price Guider\Sample Pages\Catalogue  (2).xls";

            Microsoft.Office.Interop.Excel.Application ExcelObj = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook theWorkbook = null;
            Microsoft.Office.Interop.Excel.Worksheet excelWorksheet = null;
            bool SaveChanges = false;

            try
            {
                theWorkbook = ExcelObj.Workbooks.Open(filename.ToString(), 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true);

                Microsoft.Office.Interop.Excel.Sheets sheets = theWorkbook.Worksheets;

                excelWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)sheets.get_Item(1);

                // check spreadsheet headers




                /*
                // get the area of the spreadsheet that contains the details
                Microsoft.Office.Interop.Excel.Range range = excelWorksheet.get_Range("Dep_AllDepotDetails");

                System.Array myvalues = (System.Array)range.Cells.Value;

                // loop through each depot row
                for (int rowIndex = 1; rowIndex <= range.Cells.Rows.Count; rowIndex++)
                {
                            
                    // get all details for the current depot
                    //depotToAdd.Depot_Code = (myvalues.GetValue(rowIndex, C_D_DEPOT_CODE_COL) ?? "").ToString().Trim();

                } // end for
                */


                excelWorksheet.get_Range("U1").Value = "Min";
                excelWorksheet.get_Range("V1").Value = "Max";
                excelWorksheet.get_Range("W1").Value = "Fran";
                excelWorksheet.get_Range("X1").Value = "Ind";
                excelWorksheet.get_Range("Y1").Value = "Min";
                excelWorksheet.get_Range("Z1").Value = "Max";
                excelWorksheet.get_Range("AA1").Value = "Exp";
                excelWorksheet.get_Range("AB1").Value = "P Ex";


                Microsoft.Office.Interop.Excel.Range range = excelWorksheet.get_Range("B3:Z67");

                System.Array myvalues = (System.Array)range.Cells.Value;

                for (int rowIndex = 2; rowIndex <= range.Cells.Rows.Count; rowIndex++)
                {

                    string lot = (myvalues.GetValue(rowIndex, 1) ?? "").ToString().Trim();
                    string make = (myvalues.GetValue(rowIndex, 3) ?? "").ToString().Trim();
                    string model = (myvalues.GetValue(rowIndex, 5) ?? "").ToString().Trim();
                    string bodyType = (myvalues.GetValue(rowIndex, 6) ?? "").ToString().Trim();
                    string regDate_str = (myvalues.GetValue(rowIndex, 7) ?? "").ToString().Trim();
                    string fuel = (myvalues.GetValue(rowIndex, 9) ?? "").ToString().Trim();
                    string transmission = (myvalues.GetValue(rowIndex, 10) ?? "").ToString().Trim();
                    string doors = (myvalues.GetValue(rowIndex, 12) ?? "").ToString().Trim();

                    //Console.WriteLine("Lot: " + lot);
                    //Console.WriteLine("Make: " + make);
                    //Console.WriteLine("Model: " + model);
                    //Console.WriteLine("Bodytype: " + bodyType);
                    //Console.WriteLine("Reg Date: " + regDate_str);
                    //Console.WriteLine("Fuel: " + fuel);
                    //Console.WriteLine("Trans: " + transmission);
                    //Console.WriteLine("Doors: " + doors);

                    DateTime regDate_dte;

                    if (DateTime.TryParse(regDate_str, out regDate_dte))
                    {
                        //Console.WriteLine("Conv Date: " + regDate_dte.Year);

                        string combinedLine = regDate_dte.Year + " " +
                                          make + " " +
                                          model + " " +
                                          bodyType + " " +
                                          fuel + " " +
                                          transmission;

                        Console.WriteLine(combinedLine);


                        range[rowIndex, 19].Value = combinedLine;

                        range[rowIndex, 19].ColumnWidth = 90;

                        range[rowIndex, 28].ColumnWidth = 100;

                        //////////////////////////////////
                        // NOW DO WEB PAGE READING STUFF
                        //////////////////////////////////

                        // send combined line to web page reader

                        string baseURL = @"http://www.honestjohn.co.uk/used-prices/";
                        string args = "?q=";

                        string fullURL = baseURL + args + combinedLine;

                        string fullFileName = Path.Combine(workingDir, combinedLine + ".html");

                        SaveWebPage(fullURL, fullFileName);

                        
                        // process the saved page

                        CarValuation_HonestJohn carPrices = this.ProcessFile(fullFileName);

                        /////////////////////////////////////////////
                        // NOW SAVE THE FOUND VALUES BACK TO THE SS
                        /////////////////////////////////////////////

                        range[rowIndex, 20].Value = carPrices.DealerMinPrice;
                        range[rowIndex, 21].Value = carPrices.DealerMaxPrice;
                        range[rowIndex, 22].Value = carPrices.DealerFranchisedExpectedPrice;
                        range[rowIndex, 23].Value = carPrices.DealerIndependantExpectedPrice;
                        range[rowIndex, 24].Value = carPrices.PrivateSellerMinPrice;
                        range[rowIndex, 25].Value = carPrices.PrivateSellerMaxPrice;
                        range[rowIndex, 26].Value = carPrices.PrivateSellerExpectedPrice;
                        range[rowIndex, 27].Value = carPrices.PartExPrice;
                        range[rowIndex, 28].Value = carPrices.AllPriceText;

                        //break;
                    } // end if

                    int secondsToWait = 20;

                    Thread.Sleep(secondsToWait * 1000);

                } // end for

                // Set a flag saying that all is well and it is ok to save our changes to a file.
                SaveChanges = true;

                string newFilename = @"D:\Users\David\Documents\Visual Studio 2010\Projects\Car Price Guider\Sample Pages\newCat1.xls";

                //  Save the file to disk
                theWorkbook.SaveAs(newFilename, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal,
                          null, null, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive,
                          false, false, null, null, null);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message, ex);
            }
            finally
            {

                try
                {
                    // Repeat xl.Visible and xl.UserControl releases just to be sure
                    // we didn't error out ahead of time.

                    if (ExcelObj != null)
                    {
                        ExcelObj.Visible = false;
                        ExcelObj.UserControl = false;

                        if (theWorkbook != null)
                        {
                            // Close the document and avoid user prompts to save if our method failed.
                            theWorkbook.Close(SaveChanges, null, null);
                        }

                        ExcelObj.Workbooks.Close();
                    }
                }
                catch { }

                // Gracefully exit out and destroy all COM objects to avoid hanging instances
                // of Excel.exe whether our method failed or not.

                if (ExcelObj != null) { ExcelObj.Quit(); }

                if (excelWorksheet != null) { Marshal.ReleaseComObject(excelWorksheet); }
                if (theWorkbook != null) { Marshal.ReleaseComObject(theWorkbook); }
                if (ExcelObj != null) { Marshal.ReleaseComObject(ExcelObj); }

                excelWorksheet = null;
                theWorkbook = null;
                ExcelObj = null;
                GC.Collect();
            }// end try/catch/finally

            MessageBox.Show(this, "Finished");

        }

        private void button6_Click(object sender, EventArgs e)
        {

            string workingDir = @"D:\Users\David\Documents\Visual Studio 2010\Projects\Car Price Guider\Sample Pages\2013-09-28T21-09-22";


            string fullFileName = Path.Combine(workingDir, "2003 SUBARU IMPREZA GX SPORT AWD - 1994cc 4dr Saloon Petrol Manual" + ".html");

            CarValuation_HonestJohn carPrices = this.ProcessFile1(fullFileName);

            Console.WriteLine("*****************************************");
            Console.WriteLine("Dealer Min: " + carPrices.DealerMinPrice);
            Console.WriteLine("Dealer Max: " + carPrices.DealerMaxPrice);
            Console.WriteLine("Franc Dealer Exp: " + carPrices.DealerFranchisedExpectedPrice);
            Console.WriteLine("Indep Dealer Exp: " + carPrices.DealerIndependantExpectedPrice);
            Console.WriteLine("Private Seller Min: " + carPrices.PrivateSellerMinPrice);
            Console.WriteLine("Private Seller Max: " + carPrices.PrivateSellerMaxPrice);
            Console.WriteLine("Private Seller Exp: " + carPrices.PrivateSellerExpectedPrice);
            Console.WriteLine("Part Ex Exp: " + carPrices.PartExPrice);
            Console.WriteLine(carPrices.AllPriceText);
            Console.WriteLine("*****************************************");
            fullFileName = Path.Combine(workingDir, "2003 VAUXHALL ZAFIRA CLUB 16V - 1598cc 5dr MPV Petrol Manual" + ".html");

            CarValuation_HonestJohn carPrices1 = this.ProcessFile1(fullFileName);

            Console.WriteLine("*****************************************");
            Console.WriteLine("Dealer Min: " + carPrices1.DealerMinPrice);
            Console.WriteLine("Dealer Max: " + carPrices1.DealerMaxPrice);
            Console.WriteLine("Franc Dealer Exp: " + carPrices1.DealerFranchisedExpectedPrice);
            Console.WriteLine("Indep Dealer Exp: " + carPrices1.DealerIndependantExpectedPrice);
            Console.WriteLine("Private Seller Min: " + carPrices1.PrivateSellerMinPrice);
            Console.WriteLine("Private Seller Max: " + carPrices1.PrivateSellerMaxPrice);
            Console.WriteLine("Private Seller Exp: " + carPrices1.PrivateSellerExpectedPrice);
            Console.WriteLine("Part Ex Exp: " + carPrices1.PartExPrice);
            Console.WriteLine(carPrices1.AllPriceText);
            Console.WriteLine("*****************************************");

        }

        private void btn_GetHonestJohnValuation_Click(object sender, EventArgs e)
        {
            
            string workingDir = @"D:\Users\David\Documents\Stuff To Keep Synced\Visual Studio 2010\Projects\Car Price Guider\Sample Pages\2013-09-28T21-09-22";
            workingDir = txtBox_SamplePageDir.Text;

            //string fullFileName = Path.Combine(workingDir, "2003 SUBARU IMPREZA GX SPORT AWD - 1994cc 4dr Saloon Petrol Manual" + ".html");
            string fullFileName = Path.Combine(workingDir, textBox8.Text);

            CarValuation_HonestJohn carPrices = ValuationParser_HonestJohn.GetValuation_StoredFile(fullFileName);

            Console.WriteLine("*****************************************");
            Console.WriteLine("Dealer Min: " + carPrices.DealerMinPrice);
            Console.WriteLine("Dealer Max: " + carPrices.DealerMaxPrice);
            Console.WriteLine("Dealer Exp: " + carPrices.DealerExpectedPrice);
            Console.WriteLine("Franc Dealer Exp: " + carPrices.DealerFranchisedExpectedPrice);
            Console.WriteLine("Indep Dealer Exp: " + carPrices.DealerIndependantExpectedPrice);
            Console.WriteLine("Private Seller Min: " + carPrices.PrivateSellerMinPrice);
            Console.WriteLine("Private Seller Max: " + carPrices.PrivateSellerMaxPrice);
            Console.WriteLine("Private Seller Exp: " + carPrices.PrivateSellerExpectedPrice);
            Console.WriteLine("Part Ex Exp: " + carPrices.PartExPrice);
            Console.WriteLine(carPrices.AllPriceText);
            Console.WriteLine("*****************************************");



            txtbox_HJoutput.AppendText("*****************************************\n");
            txtbox_HJoutput.AppendText("Dealer Min: " + carPrices.DealerMinPrice + "\n");
            txtbox_HJoutput.AppendText("Dealer Max: " + carPrices.DealerMaxPrice + "\n");
            txtbox_HJoutput.AppendText("Dealer Exp: " + carPrices.DealerExpectedPrice + "\n");
            txtbox_HJoutput.AppendText("Franc Dealer Exp: " + carPrices.DealerFranchisedExpectedPrice + "\n");
            txtbox_HJoutput.AppendText("Indep Dealer Exp: " + carPrices.DealerIndependantExpectedPrice + "\n");
            txtbox_HJoutput.AppendText("Private Seller Min: " + carPrices.PrivateSellerMinPrice + "\n");
            txtbox_HJoutput.AppendText("Private Seller Max: " + carPrices.PrivateSellerMaxPrice + "\n");
            txtbox_HJoutput.AppendText("Private Seller Exp: " + carPrices.PrivateSellerExpectedPrice + "\n");
            txtbox_HJoutput.AppendText("Part Ex Exp: " + carPrices.PartExPrice + "\n");
            txtbox_HJoutput.AppendText(carPrices.AllPriceText + "\n");
            txtbox_HJoutput.AppendText("*****************************************\n");

            return;

            fullFileName = Path.Combine(workingDir, "2003 VAUXHALL ZAFIRA CLUB 16V - 1598cc 5dr MPV Petrol Manual" + ".html");

            //CarValuation carPrices1 = this.ProcessFile1(fullFileName);
            CarValuation_HonestJohn carPrices1 = ValuationParser_HonestJohn.GetValuation_StoredFile(fullFileName);

            Console.WriteLine("*****************************************");
            Console.WriteLine("Dealer Min: " + carPrices1.DealerMinPrice);
            Console.WriteLine("Dealer Max: " + carPrices1.DealerMaxPrice);
            Console.WriteLine("Dealer Exp: " + carPrices1.DealerExpectedPrice + "\n");
            Console.WriteLine("Franc Dealer Exp: " + carPrices1.DealerFranchisedExpectedPrice);
            Console.WriteLine("Indep Dealer Exp: " + carPrices1.DealerIndependantExpectedPrice);
            Console.WriteLine("Private Seller Min: " + carPrices1.PrivateSellerMinPrice);
            Console.WriteLine("Private Seller Max: " + carPrices1.PrivateSellerMaxPrice);
            Console.WriteLine("Private Seller Exp: " + carPrices1.PrivateSellerExpectedPrice);
            Console.WriteLine("Part Ex Exp: " + carPrices1.PartExPrice);
            Console.WriteLine(carPrices1.AllPriceText);
            Console.WriteLine("*****************************************");

            txtbox_HJoutput.AppendText("*****************************************\n");
            txtbox_HJoutput.AppendText("Dealer Min: " + carPrices1.DealerMinPrice + "\n");
            txtbox_HJoutput.AppendText("Dealer Max: " + carPrices1.DealerMaxPrice + "\n");
            txtbox_HJoutput.AppendText("Dealer Exp: " + carPrices1.DealerExpectedPrice + "\n");
            txtbox_HJoutput.AppendText("Franc Dealer Exp: " + carPrices1.DealerFranchisedExpectedPrice + "\n");
            txtbox_HJoutput.AppendText("Indep Dealer Exp: " + carPrices1.DealerIndependantExpectedPrice + "\n");
            txtbox_HJoutput.AppendText("Private Seller Min: " + carPrices1.PrivateSellerMinPrice + "\n");
            txtbox_HJoutput.AppendText("Private Seller Max: " + carPrices1.PrivateSellerMaxPrice + "\n");
            txtbox_HJoutput.AppendText("Private Seller Exp: " + carPrices1.PrivateSellerExpectedPrice + "\n");
            txtbox_HJoutput.AppendText("Part Ex Exp: " + carPrices1.PartExPrice + "\n");
            txtbox_HJoutput.AppendText(carPrices1.AllPriceText + "\n");
            txtbox_HJoutput.AppendText("*****************************************\n");

        }










        #region OLD


        public string GetPathText(HtmlAgilityPack.HtmlDocument htmlDoc, string xPath)
        {
            var nodes = htmlDoc.DocumentNode.SelectNodes(xPath);
            //src = new List<string>(nodes.Count);

            if (nodes != null)
            {
                foreach (var node in nodes)
                {
                    if (node.Id != null)
                    {
                        //src.Add(node.Id);
                        //Console.WriteLine("Most adverts are between " + node.InnerText);
                        string textFound = node.InnerText.Trim();
                        textFound = textFound.Replace("Â", "");
                        return textFound;
                    }

                }
            }

            return null;
        }


        private CarValuation_HonestJohn ProcessFile1(string fname)
        {

            CarValuation_HonestJohn returnPrices = new CarValuation_HonestJohn();

            var req1 = (FileWebRequest)WebRequest.Create(fname);
            req1.Method = "GET";

            using (WebResponse odpoved = req1.GetResponse())
            {
                var htmlDoc = new HtmlAgilityPack.HtmlDocument();
                htmlDoc.Load(odpoved.GetResponseStream());

                /*
                 * Below are working
                //*[@class="pricePoints"]/tr[1]/td[1]/b[1] - dealer
                //*[@class="pricePoints"]/tr[1]/td[3]/b[1] - 4050
                //*[@class="pricePoints"]/tr[1]/td[3]/b[2] - 5200
                 * 
                 */

                string minPrice = GetPathText(htmlDoc, "//*[@id=\"prices\"]/div[3]/div/div/div[3]/p[1]/b[1]");
                string maxPrice = GetPathText(htmlDoc, "//*[@id=\"prices\"]/div[3]/div/div/div[3]/p[1]/b[2]");

                string typeName1 = GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]/tr[1]/td[1]/b[1]");
                string typeName2 = GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]/tr[2]/td[1]/b[1]");
                string typeName3 = GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]/tr[3]/td[1]/b[1]");
                string typeName4 = GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]/tr[4]/td[1]/b[1]");
                string typeName5 = GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]/tr[5]/td[1]/b[1]");


                string type1MinPrice = GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]/tr[1]/td[3]/b[1]");
                string type1MaxPrice = GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]/tr[1]/td[3]/b[2]");
                string type1ExpPrice = GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]/tr[2]/td[3]/b[1]");

                string type2MinPrice = GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]/tr[2]/td[3]/b[1]");
                string type2MaxPrice = GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]/tr[2]/td[3]/b[2]");
                string type2ExpPrice = GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]/tr[2]/td[3]/b[1]");

                string type3MinPrice = GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]/tr[3]/td[3]/b[1]");
                string type3MaxPrice = GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]/tr[3]/td[3]/b[2]");
                string type3ExpPrice = GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]/tr[3]/td[3]/b[1]");

                string type4MinPrice = GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]/tr[4]/td[3]/b[1]");
                string type4MaxPrice = GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]/tr[4]/td[3]/b[2]");

                string type5ExpectedPrice = GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]/tr[5]/td[2]/b[1]");

                string partExPrice = GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]/tr[6]/td[3]/b[1]");

                Console.WriteLine("Most adverts are between " + minPrice + " and " + maxPrice);


                Console.WriteLine("Type1: " + typeName1);
                Console.WriteLine("Type2: " + typeName2);
                Console.WriteLine("Type3: " + typeName3);
                Console.WriteLine("Type4: " + typeName4);
                Console.WriteLine("Type5: " + typeName5);

                if (typeName1 == "dealer")
                {
                    returnPrices.DealerMinPrice = type1MinPrice;
                    returnPrices.DealerMaxPrice = type1MaxPrice;
                }
                else if (typeName1 == "franchised dealer")
                {
                    returnPrices.DealerFranchisedExpectedPrice = type1ExpPrice;
                }
                else if (typeName1 == "private seller")
                {
                    returnPrices.PrivateSellerMinPrice = type1MinPrice;
                    returnPrices.PrivateSellerMaxPrice = type1MaxPrice;
                }


                if (typeName2 == "dealer")
                {
                    returnPrices.DealerMinPrice = type2MinPrice;
                    returnPrices.DealerMaxPrice = type2MaxPrice;
                }
                else if (typeName2 == "franchised dealer")
                {
                    returnPrices.DealerFranchisedExpectedPrice = type2ExpPrice;
                }
                else if (typeName2 == "private seller")
                {
                    returnPrices.PrivateSellerMinPrice = type2MinPrice;
                    returnPrices.PrivateSellerMaxPrice = type2MaxPrice;
                }


                if (typeName3 == "dealer")
                {
                    returnPrices.DealerMinPrice = type3MinPrice;
                    returnPrices.DealerMaxPrice = type3MaxPrice;
                }
                else if (typeName3 == "franchised dealer")
                {
                    returnPrices.DealerFranchisedExpectedPrice = type3ExpPrice;
                }
                else if (typeName3 == "private seller")
                {
                    returnPrices.PrivateSellerMinPrice = type3MinPrice;
                    returnPrices.PrivateSellerMaxPrice = type3MaxPrice;
                }

                //returnPrices.DealerFranchisedExpectedPrice = dealerFranchisedExpectedPrice;
                //returnPrices.DealerIndependantExpectedPrice = dealerIndependantExpectedPrice;



                //returnPrices.PrivateSellerExpectedPrice = privateSellerExpectedPrice;

                returnPrices.PartExPrice = partExPrice;



                // All this below will format the text in 1 block
                string allPriceText = GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]").Trim(); // this works atm

                allPriceText = allPriceText.Trim();

                if (allPriceText != "")
                {
                    allPriceText = allPriceText.Replace("\r\n", " ");
                    allPriceText = allPriceText.Replace("\r", " ");
                    allPriceText = allPriceText.Replace("\n", " ");
                    allPriceText = allPriceText.Replace("\t", " ");
                    allPriceText = allPriceText.Replace("At", "\nAt");
                    allPriceText = allPriceText.Replace("From", "\nFrom");
                    allPriceText = allPriceText.Replace("If selling", "\nIf selling");

                    for (int i = 0; i < 200; i++)
                    {
                        allPriceText = allPriceText.Replace("  ", " ");
                    }

                    allPriceText = allPriceText.Replace("Â", " ");
                    allPriceText = allPriceText.Replace("&nbsp;", " ");

                    returnPrices.AllPriceText = allPriceText;

                    //Console.WriteLine(minPrice);
                }

            } // end using

            return returnPrices;

        }


        #endregion



        private void button7_Click(object sender, EventArgs e)
        {

            string workingDir = @"D:\Users\David\Documents\Stuff To Keep Synced\Visual Studio 2010\Projects\Car Price Guider\Sample Pages\";
            workingDir = txtBox_SamplePageDir.Text;

            //string fullFileName = Path.Combine(workingDir, "2003 SUBARU IMPREZA GX SPORT AWD - 1994cc 4dr Saloon Petrol Manual" + ".html");
            //textBox8.Text = @"Ford Fiesta 2009 1.6 Zetec S Price Guide   Honest John.htm";
            string fullFileName = Path.Combine(workingDir, textBox8.Text);

            CarValuation_HonestJohn carPrices = ValuationParser_HonestJohn.GetValuation_StoredFile1(fullFileName);

            //Console.WriteLine("*****************************************");
            //Console.WriteLine("Dealer Min: " + carPrices.DealerMinPrice);
            //Console.WriteLine("Dealer Max: " + carPrices.DealerMaxPrice);
            //Console.WriteLine("Dealer Exp: " + carPrices.DealerExpectedPrice);
            //Console.WriteLine("Franc Dealer Exp: " + carPrices.DealerFranchisedExpectedPrice);
            //Console.WriteLine("Indep Dealer Exp: " + carPrices.DealerIndependantExpectedPrice);
            //Console.WriteLine("Private Seller Min: " + carPrices.PrivateSellerMinPrice);
            //Console.WriteLine("Private Seller Max: " + carPrices.PrivateSellerMaxPrice);
            //Console.WriteLine("Private Seller Exp: " + carPrices.PrivateSellerExpectedPrice);
            //Console.WriteLine("Part Ex Exp: " + carPrices.PartExPrice);
            //Console.WriteLine(carPrices.AllPriceText);
            //Console.WriteLine("*****************************************");



            txtbox_HJoutput.AppendText("*****************************************\n");
            txtbox_HJoutput.AppendText("Dealer Min: " + carPrices.DealerMinPrice + "\n");
            txtbox_HJoutput.AppendText("Dealer Max: " + carPrices.DealerMaxPrice + "\n");
            txtbox_HJoutput.AppendText("Dealer Exp: " + carPrices.DealerExpectedPrice + "\n");
            txtbox_HJoutput.AppendText("Franc Dealer Exp: " + carPrices.DealerFranchisedExpectedPrice + "\n");
            txtbox_HJoutput.AppendText("Indep Dealer Exp: " + carPrices.DealerIndependantExpectedPrice + "\n");
            txtbox_HJoutput.AppendText("Car supermarket Exp: " + carPrices.CarSupermarketExpectedPrice + "\n");
            txtbox_HJoutput.AppendText("Private Seller Min: " + carPrices.PrivateSellerMinPrice + "\n");
            txtbox_HJoutput.AppendText("Private Seller Max: " + carPrices.PrivateSellerMaxPrice + "\n");
            txtbox_HJoutput.AppendText("Private Seller Exp: " + carPrices.PrivateSellerExpectedPrice + "\n");
            txtbox_HJoutput.AppendText("Part Ex Exp: " + carPrices.PartExPrice + "\n");
            txtbox_HJoutput.AppendText(carPrices.AllPriceText + "\n");
            txtbox_HJoutput.AppendText("*****************************************\n");

            //string readable_text = carPrices.AllPriceText;

            //var result = Regex.Split(readable_text, "\r\n|\r|\n");

            ///*
            // * 
            // * At a dealer, aim to pay between £7,350 and £9,300 
            //    At a franchised dealer, expect to pay £8,550. 
            //    At an independent dealer, expect to pay £8,000. 
            //    At a car supermarket, expect to pay £7,900. 
            //    From a private seller, aim to pay between  £6,950 and £8,150 and expect to pay £7,500. 
            //    If selling in part exchange, expect to receive £6,530.
            // * 
            // * 
            //*/


            //txtbox_HJoutput.AppendText("*****************************************\n");
            //txtbox_HJoutput.AppendText("*****************************************\n");


            //CarValuation valuation = new CarValuation();


            //for (int i = 0; i < result.Length; i++)
            //{
            //    string currLine = result[i];
            //    currLine = currLine.Replace("At an", "");
            //    currLine = currLine.Replace("At a", "");
            //    currLine = currLine.Replace("From a", "");
            //    currLine = currLine.Replace("If selling in", "");
            //    Match regex_Res_SellerType = Regex.Match(currLine, @"^[a-zA-Z ]*");
            //    string sellerType = regex_Res_SellerType.Value;

            //    txtbox_HJoutput.AppendText("Seller Type: " + sellerType + "\n");

            //    if (sellerType != "")
            //    {
            //        MatchCollection prices = Regex.Matches(currLine, @"£[0-9,]*");

            //        switch (sellerType.ToUpper().Trim())
            //        {
            //            case "DEALER":
            //                valuation.DealerMinPrice = prices[0].Value;
            //                valuation.DealerMaxPrice = prices[1].Value;
            //                if (prices.Count == 3)
            //                {
            //                    valuation.DealerExpectedPrice = prices[2].Value;
            //                }
            //                break;
            //            case "FRANCHISED DEALER":
            //                valuation.DealerFranchisedExpectedPrice = prices[0].Value;
            //                break;
            //            case "INDEPENDENT DEALER":
            //                valuation.DealerIndependantExpectedPrice = prices[0].Value;
            //                break;
            //            case "CAR SUPERMARKET":
            //                valuation.CarSupermarketExpectedPrice = prices[0].Value;
            //                break;
            //            case "PRIVATE SELLER":
            //                valuation.PrivateSellerMinPrice = prices[0].Value;
            //                valuation.PrivateSellerMaxPrice = prices[1].Value;
            //                if (prices.Count == 3)
            //                {
            //                    valuation.PrivateSellerExpectedPrice = prices[2].Value;
            //                }
            //                break;
            //            case "PART EXCHANGE":
            //                valuation.PartExPrice = prices[0].Value;
            //                break;
            //            default:
            //                break;
            //        } // end switch
            //    } // end if

            //} // end for

            //txtbox_HJoutput.AppendText("*****************************************\n");
            //txtbox_HJoutput.AppendText("*****************************************\n");

            //txtbox_HJoutput.AppendText("*****************************************\n");
            //txtbox_HJoutput.AppendText("Dealer Min: " + valuation.DealerMinPrice + "\n");
            //txtbox_HJoutput.AppendText("Dealer Max: " + valuation.DealerMaxPrice + "\n");
            //txtbox_HJoutput.AppendText("Dealer Exp: " + valuation.DealerExpectedPrice + "\n");
            //txtbox_HJoutput.AppendText("Franc Dealer Exp: " + valuation.DealerFranchisedExpectedPrice + "\n");
            //txtbox_HJoutput.AppendText("Indep Dealer Exp: " + valuation.DealerIndependantExpectedPrice + "\n");
            //txtbox_HJoutput.AppendText("Private Seller Min: " + valuation.PrivateSellerMinPrice + "\n");
            //txtbox_HJoutput.AppendText("Private Seller Max: " + valuation.PrivateSellerMaxPrice + "\n");
            //txtbox_HJoutput.AppendText("Private Seller Exp: " + valuation.PrivateSellerExpectedPrice + "\n");
            //txtbox_HJoutput.AppendText("Part Ex Exp: " + valuation.PartExPrice + "\n");
            //txtbox_HJoutput.AppendText(valuation.AllPriceText + "\n");
            //txtbox_HJoutput.AppendText("*****************************************\n");


        }

        /// <summary>
        /// Process all the car details from a saved catalogue from Bawtry
        /// Stores the car details in memory
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_ProcessCatalogue_Bawtry_Click(object sender, EventArgs e)
        {
            carDetailsToProcess.Clear();

            string workingDir = @"D:\Users\David\Documents\Stuff To Keep Synced\Visual Studio 2010\Projects\Car Price Guider(2)\Sample Pages\";
            //workingDir = txtBox_SamplePageDir.Text;

            string fullFileName = Path.Combine(workingDir, "Wednesday Auction, 22 05 2013   Bawtry Motor Auction Remarketing.htm");
            //textBox8.Text = @"Ford Fiesta 2009 1.6 Zetec S Price Guide   Honest John.htm";
            //string fullFileName = Path.Combine(workingDir, textBox8.Text);

            openFileDialog1.Title = "Choose saved Bawtry web page";
            openFileDialog1.InitialDirectory = ".";
            openFileDialog1.FileName = "";
            openFileDialog1.Filter = "Web Page|*.htm;*.html";
            openFileDialog1.Multiselect = false;

            DialogResult dr = openFileDialog1.ShowDialog(this);
            if (dr != DialogResult.OK)
            {
                return;
            } // end if

            fullFileName = openFileDialog1.FileName;

            List<CarDetails> carDetails = AuctionCatalogueParser_Bawtry.ParseCatalogue_StoredFile(fullFileName);
            carDetailsToProcess.AddRange(carDetails.ToArray());

            foreach (var currCar in carDetailsToProcess)
            {
                textBox_CarCatalogueResults_Bawtry.AppendText("*****************************************" + System.Environment.NewLine);
                foreach (var prop in currCar.GetType().GetProperties())
                {
                    textBox_CarCatalogueResults_Bawtry.AppendText(String.Format("{0} = {1}", prop.Name, prop.GetValue(currCar, null)) + System.Environment.NewLine);
                } // end foreach
                textBox_CarCatalogueResults_Bawtry.AppendText("Mileage = " + currCar.GetMileage() + System.Environment.NewLine);
                textBox_CarCatalogueResults_Bawtry.AppendText(currCar.FormatForValuation_CAP_Email() + System.Environment.NewLine);
                textBox_CarCatalogueResults_Bawtry.AppendText("*****************************************" + System.Environment.NewLine);
            } // end foreach

        }

        List<CarDetails> carDetailsToProcess = new List<CarDetails>();

        /// <summary>
        /// Processes a list of car details that is stored in memory
        /// Requests the valuation from website and saves it to the selected folder
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button9_Click(object sender, EventArgs e)
        {

            if (carDetailsToProcess == null || carDetailsToProcess.Count == 0)
            {
                MessageBox.Show("No car details loaded");
                return;
            } // end if

            string workingDir = @"D:\Users\David\Documents\Stuff To Keep Synced\Visual Studio 2010\Projects\Car Price Guider\Sample Pages\";
            workingDir = txtBox_SamplePageDir.Text;


            DateTime startDateTime = DateTime.Now;

            //string workingDir = @"D:\Users\David\Documents\Visual Studio 2010\Projects\Car Price Guider\Sample Pages";

            string workingDirName = "";

            workingDirName = startDateTime.ToString("yyyy-MM-ddTHH-mm-ss");


            workingDir = Path.Combine(workingDir, workingDirName);

            if (Directory.Exists(workingDir))
            {
                MessageBox.Show("Directory exists");
                return;
            }
            else
            {
                Directory.CreateDirectory(workingDir);
            } // end if-then-else




            foreach (var currCar in carDetailsToProcess)
            {


                //string fullFileName = Path.Combine(workingDir, "2003 SUBARU IMPREZA GX SPORT AWD - 1994cc 4dr Saloon Petrol Manual" + ".html");
                //textBox8.Text = @"Ford Fiesta 2009 1.6 Zetec S Price Guide   Honest John.htm";
                //string fullFileName = Path.Combine(workingDir, textBox8.Text);

                string fullFileName = currCar.FormatForValuation_HonestJohn() + ".html";
                fullFileName = Path.Combine(workingDir, fullFileName);

                ValuationParser_HonestJohn.RequestValuationFromWebsite(currCar, fullFileName);

                Thread.Sleep(20000); // pause 10 seconds
                //return;
            } // end foreach

            return;

            string fullFileName1 = "";
            CarValuation_HonestJohn carPrices = ValuationParser_HonestJohn.GetValuation_StoredFile(fullFileName1);


            txtbox_HJoutput.AppendText("*****************************************\n");
            txtbox_HJoutput.AppendText("Dealer Min: " + carPrices.DealerMinPrice + "\n");
            txtbox_HJoutput.AppendText("Dealer Max: " + carPrices.DealerMaxPrice + "\n");
            txtbox_HJoutput.AppendText("Dealer Exp: " + carPrices.DealerExpectedPrice + "\n");
            txtbox_HJoutput.AppendText("Franc Dealer Exp: " + carPrices.DealerFranchisedExpectedPrice + "\n");
            txtbox_HJoutput.AppendText("Indep Dealer Exp: " + carPrices.DealerIndependantExpectedPrice + "\n");
            txtbox_HJoutput.AppendText("Car supermarket Exp: " + carPrices.CarSupermarketExpectedPrice + "\n");
            txtbox_HJoutput.AppendText("Private Seller Min: " + carPrices.PrivateSellerMinPrice + "\n");
            txtbox_HJoutput.AppendText("Private Seller Max: " + carPrices.PrivateSellerMaxPrice + "\n");
            txtbox_HJoutput.AppendText("Private Seller Exp: " + carPrices.PrivateSellerExpectedPrice + "\n");
            txtbox_HJoutput.AppendText("Part Ex Exp: " + carPrices.PartExPrice + "\n");
            txtbox_HJoutput.AppendText(carPrices.AllPriceText + "\n");
            txtbox_HJoutput.AppendText("*****************************************\n");
            
        }

        /// <summary>
        /// Processes a folder that contains saved car valuations
        /// Stores the valuations in memory
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_ProcessSavedValuations_Click(object sender, EventArgs e)
        {

            string workingDir = @"D:\Users\David\Documents\Stuff To Keep Synced\Visual Studio 2010\Projects\Car Price Guider\Sample Pages\";
            workingDir = txtBox_SamplePageDir.Text;


            var fsd = new FolderSelectDialog();
            fsd.Title = "What to select";
            fsd.InitialDirectory = workingDir;
            if (fsd.ShowDialog(IntPtr.Zero))
            {
                //Console.WriteLine(fsd.FileName);
                workingDir = fsd.FileName;
            }
            else
            {
                return;
            } // end if-then-else


            //DialogResult d_res = folderBrowserDialog1.ShowDialog();
            //DateTime startDateTime = DateTime.Now;

            //string workingDir = @"D:\Users\David\Documents\Visual Studio 2010\Projects\Car Price Guider\Sample Pages";

            //string workingDirName = "";

            //workingDirName = startDateTime.ToString("yyyy-MM-ddTHH-mm-ss");


            //workingDir = Path.Combine(workingDir, workingDirName);

            //if (Directory.Exists(workingDir))
            //{
            //    MessageBox.Show("Directory exists");
            //    return;
            //}
            //else
            //{
            //    Directory.CreateDirectory(workingDir);
            //}

            foreach (var currFile in Directory.GetFiles(workingDir, "*.html"))
            {
                //string fullFileName1 = "";
                CarValuation_HonestJohn carPrices = ValuationParser_HonestJohn.GetValuation_StoredFile(currFile);

                txtbox_HJoutput.AppendText("*****************************************" + System.Environment.NewLine);
                txtbox_HJoutput.AppendText("Car: " + Path.GetFileName(currFile) + System.Environment.NewLine);
                txtbox_HJoutput.AppendText("Dealer Min: " + carPrices.DealerMinPrice + System.Environment.NewLine);
                txtbox_HJoutput.AppendText("Dealer Max: " + carPrices.DealerMaxPrice + System.Environment.NewLine);
                txtbox_HJoutput.AppendText("Dealer Exp: " + carPrices.DealerExpectedPrice + System.Environment.NewLine);
                txtbox_HJoutput.AppendText("Franc Dealer Exp: " + carPrices.DealerFranchisedExpectedPrice + System.Environment.NewLine);
                txtbox_HJoutput.AppendText("Indep Dealer Exp: " + carPrices.DealerIndependantExpectedPrice + System.Environment.NewLine);
                txtbox_HJoutput.AppendText("Car supermarket Exp: " + carPrices.CarSupermarketExpectedPrice + System.Environment.NewLine);
                txtbox_HJoutput.AppendText("Private Seller Min: " + carPrices.PrivateSellerMinPrice + System.Environment.NewLine);
                txtbox_HJoutput.AppendText("Private Seller Max: " + carPrices.PrivateSellerMaxPrice + System.Environment.NewLine);
                txtbox_HJoutput.AppendText("Private Seller Exp: " + carPrices.PrivateSellerExpectedPrice + System.Environment.NewLine);
                txtbox_HJoutput.AppendText("Part Ex Exp: " + carPrices.PartExPrice + System.Environment.NewLine);
                txtbox_HJoutput.AppendText(carPrices.AllPriceText + System.Environment.NewLine);
                txtbox_HJoutput.AppendText("*****************************************" + System.Environment.NewLine);
            }
        }











        private void button9_Click_1(object sender, EventArgs e)
        {
            string workingDir = @"D:\Users\David\Documents\Stuff To Keep Synced\Visual Studio 2010\Projects\Car Price Guider(2)\Sample Pages\";
            //workingDir = txtBox_SamplePageDir.Text;

            string fullFileName = Path.Combine(workingDir, "Wednesday Auction, 22 05 2013   Bawtry Motor Auction Remarketing.htm");
            //textBox8.Text = @"Ford Fiesta 2009 1.6 Zetec S Price Guide   Honest John.htm";
            //string fullFileName = Path.Combine(workingDir, textBox8.Text);

            openFileDialog1.Title = "Choose saved Bawtry export file";
            openFileDialog1.InitialDirectory = ".";
            openFileDialog1.FileName = "";
            openFileDialog1.Multiselect = false;

            DialogResult dr = openFileDialog1.ShowDialog(this);
            if (dr != DialogResult.OK)
            {
                return;
            }

            fullFileName = openFileDialog1.FileName;

            List<CarDetails> carDetails = AuctionCatalogueParser_Bawtry.ParseCatalogue_StoredFile_PATHTESTER(fullFileName, textBox11.Text);
            carDetailsToProcess.AddRange(carDetails.ToArray());

            foreach (var currCar in carDetailsToProcess)
            {
                textBox_CarCatalogueResults_Bawtry.AppendText("*****************************************" + System.Environment.NewLine);
                foreach (var prop in currCar.GetType().GetProperties())
                {
                    textBox_CarCatalogueResults_Bawtry.AppendText(String.Format("{0}={1}", prop.Name, prop.GetValue(currCar, null)) + System.Environment.NewLine);
                } // end foreach
                textBox_CarCatalogueResults_Bawtry.AppendText("*****************************************" + System.Environment.NewLine);
            } // end foreach
        }

        private void btn_ProcessCatalogue_Newark_Click(object sender, EventArgs e)
        {
            carDetailsToProcess.Clear();

            openFileDialog1.Title = "Choose saved Newark export file";
            openFileDialog1.InitialDirectory = ".";
            openFileDialog1.FileName = "";
            openFileDialog1.Multiselect = false;

            DialogResult dr = openFileDialog1.ShowDialog(this);
            if (dr != DialogResult.OK)
            {
                return;
            } // end if

            string fullFileName = openFileDialog1.FileName;

            List<CarDetails> carDetails = AuctionCatalogueParser_Newark.ParseCatalogue_StoredFile(fullFileName);

            //List<CarDetails> carDetails = AuctionCatalogueParser_Bawtry.ParseCatalogue_StoredFile(fullFileName);
            carDetailsToProcess.AddRange(carDetails.ToArray());

            foreach (var currCar in carDetailsToProcess)
            {
                textBox_CarCatalogueResults_Newark.AppendText("*****************************************" + System.Environment.NewLine);
                foreach (var prop in currCar.GetType().GetProperties())
                {
                    textBox_CarCatalogueResults_Newark.AppendText(String.Format("{0} = {1}", prop.Name, prop.GetValue(currCar, null)) + System.Environment.NewLine);
                } // end foreach
                //textBox_CarCatalogueResults_Newark.AppendText("Long Description = " + currCar.Long_Description + System.Environment.NewLine);
                //textBox_CarCatalogueResults_Newark.AppendText("Mileage = " + currCar.GetMileage() + System.Environment.NewLine);
                textBox_CarCatalogueResults_Newark.AppendText(currCar.FormatForValuation_CAP_Email() + System.Environment.NewLine);
                textBox_CarCatalogueResults_Newark.AppendText("*****************************************" + System.Environment.NewLine);
            } // end foreach
        }

        private void button8_Click(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show(this, "This will create all emails and send automatically - are you sure?", "", MessageBoxButtons.OKCancel);
            if (dr == System.Windows.Forms.DialogResult.No)
            {
                return;
            }

            int emailCounter = 0;

            Microsoft.Office.Interop.Outlook.Application outlookApp = new Microsoft.Office.Interop.Outlook.Application();

            foreach (var currCar in carDetailsToProcess)
            {
                Microsoft.Office.Interop.Outlook.MailItem mailItem = (Microsoft.Office.Interop.Outlook.MailItem)
                outlookApp.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);
                mailItem.Subject = currCar.FormatForValuation_CAP_Email();
                mailItem.To = "value@cap.co.uk";
                mailItem.Body = currCar.Lot_Number + Environment.NewLine + currCar.Long_Description;
                mailItem.Display(false);
                Thread.Sleep(30000);
                mailItem.Send();
                emailCounter++;
            } // end foreach

            MessageBox.Show("Created " + emailCounter + " emails");

        }

        private void button10_Click(object sender, EventArgs e)
        {
            List<CarDetails> carValuations = new List<CarDetails>();

            openFileDialog1.Title = "Choose saved CAP Valuations";
            openFileDialog1.InitialDirectory = ".";
            openFileDialog1.FileName = "";
            openFileDialog1.Filter = "CSV|*.csv";
            openFileDialog1.Multiselect = false;

            DialogResult dr = openFileDialog1.ShowDialog(this);
            if (dr != DialogResult.OK)
            {
                return;
            }

            string fullFileName_CAPValuations = openFileDialog1.FileName;

            int lineNo = 1;
            int col_num_vrm = -1;
            int col_num_live_clean = -1;
            int col_num_live_ave = -1;
            int col_num_live_below = -1;
            int col_num_live_retail = -1;

            TextFieldParser parser = new TextFieldParser(fullFileName_CAPValuations);
            parser.TextFieldType = FieldType.Delimited;
            parser.SetDelimiters(",");
            while (!parser.EndOfData)
            {
                //Processing row
                string[] fields = parser.ReadFields();

                if (lineNo == 1) // header line
                {
                    for (int i = 0; i < fields.Length; i++)
                    {
                        if (String.Compare(fields[i], "VRM", true) == 0)
                        {
                            col_num_vrm = i;
                        } else if (String.Compare(fields[i], "Live Clean Value", true) == 0)
                        {
                            col_num_live_clean = i;
                        }
                        else if (String.Compare(fields[i], "Live Average Value", true) == 0)
                        {
                            col_num_live_ave = i;
                        }
                        else if (String.Compare(fields[i], "Live Below Value", true) == 0)
                        {
                            col_num_live_below = i;
                        }
                        else if (String.Compare(fields[i], "Live Retail", true) == 0)
                        {
                            col_num_live_retail = i;
                        } // end if-then-else
                    } // end foreach


                    if (col_num_vrm == -1)
                    {
                        throw new Exception("Column not found in CSV");
                    } // end if

                    if (col_num_live_clean == -1)
                    {
                        throw new Exception("Column not found in CSV");
                    } // end if

                    if (col_num_live_ave == -1)
                    {
                        throw new Exception("Column not found in CSV");
                    } // end if

                    if (col_num_live_below == -1)
                    {
                        throw new Exception("Column not found in CSV");
                    } // end if

                    if (col_num_live_retail == -1)
                    {
                        throw new Exception("Column not found in CSV");
                    } // end if

                }
                else
                {
                    string vrm = fields[col_num_vrm];

                    CarValuation_CAP carValuation = new CarValuation_CAP();
                    
                    double currVal;

                    if (Double.TryParse(fields[col_num_live_clean], out currVal))
                    {
                        carValuation.LivePrice_Clean = currVal;
                    } // end if

                    if (Double.TryParse(fields[col_num_live_ave], out currVal))
                    {
                        carValuation.LivePrice_Ave = currVal;
                    } // end if

                    if (Double.TryParse(fields[col_num_live_below], out currVal))
                    {
                        carValuation.LivePrice_Below = currVal;
                    } // end if

                    if (Double.TryParse(fields[col_num_live_retail], out currVal))
                    {
                        carValuation.LivePrice_Retail = currVal;
                    } // end if



                    // now we need to find this car in the car details list
                    
                    // check there's only 1 depot with this code in the spreadsheet
                    List<CarDetails> results = carDetailsToProcess.FindAll(
                    delegate(CarDetails carDets) { return carDets.RegNo == vrm; });
                    if (results.Count == 0) // no results found                
                    {
                    }
                    else if (results.Count > 1)
                    {
                    }
                    else
                    {
                        results[0].CarValuation_Cap = carValuation;
                        carValuations.Add(results[0]);
                    } // end if-then-else

                } // end if-then-else

                lineNo++;
            }
            parser.Close();


            textBox_CarCatalogueResults_Bawtry.AppendText("*****************************************" + System.Environment.NewLine);
            textBox_CarCatalogueResults_Bawtry.AppendText("*****************************************" + System.Environment.NewLine);
            textBox_CarCatalogueResults_Bawtry.AppendText("*****************************************" + System.Environment.NewLine);

            foreach (var currCar in carValuations)
            {
                textBox_CarCatalogueResults_Bawtry.AppendText("*****************************************" + System.Environment.NewLine);
                foreach (var prop in currCar.GetType().GetProperties())
                {
                    textBox_CarCatalogueResults_Bawtry.AppendText(String.Format("{0} = {1}", prop.Name, prop.GetValue(currCar, null)) + System.Environment.NewLine);
                } // end foreach
                textBox_CarCatalogueResults_Bawtry.AppendText("Mileage = " + currCar.GetMileage() + System.Environment.NewLine);
                textBox_CarCatalogueResults_Bawtry.AppendText(currCar.FormatForValuation_CAP_Email() + System.Environment.NewLine);

                CarValuation_CAP carValuation = currCar.CarValuation_Cap;
                
                //textBox_CarCatalogueResults_Bawtry.AppendText("*****************************************" + System.Environment.NewLine);
                foreach (var prop in carValuation.GetType().GetProperties())
                {
                    textBox_CarCatalogueResults_Bawtry.AppendText(String.Format("{0} = {1}", prop.Name, prop.GetValue(carValuation, null)) + System.Environment.NewLine);
                } // end foreach
                
                textBox_CarCatalogueResults_Bawtry.AppendText("*****************************************" + System.Environment.NewLine);
            } // end foreach

            this.Refresh();

            saveFileDialog1.Title = "Choose Valuation Report Location";
            saveFileDialog1.InitialDirectory = ".";
            saveFileDialog1.FileName = "";
            saveFileDialog1.Filter = "Excel|*.xls";
            dr = saveFileDialog1.ShowDialog(this);
            if (dr != DialogResult.OK)
            {
                return;
            } // end if

            string fullFileName_ValuationsReport = saveFileDialog1.FileName;

            FormatValuationsForExport(carValuations, fullFileName_ValuationsReport);

            MessageBox.Show("Finished");

        }

        private void FormatValuationsForExport(List<CarDetails> CarValuations, string fullFileName_ValuationsReport)
        {
            const int COL_FIRST_COL_NUM = 1;
            const int ROW_FIRST_ROW_NUM = 0;

            const int COL_NUM_LOT_NO = 1;
            const int ROW_NUM_LOT_NO = 0;
            
            const int COL_NUM_DESC = 2;
            const int ROW_NUM_DESC = 0;

            const int COL_NUM_REGNO = 2;
            const int ROW_NUM_REGNO = 1;

            const int COL_NUM_VALUATIONS = 3;
            const int ROW_NUM_VALUE_CLEAN = 1;
            const int ROW_NUM_VALUE_AVE = 2;
            const int ROW_NUM_VALUE_BELOW = 3;
            
            const int COL_NUM_VALUE_RETAIL = 4;
            const int ROW_NUM_VALUE_RETAIL = 1;

            const int COL_NUM_PROFIT = 5;
            const int ROW_NUM_PROFIT_CLEAN = 1;
            const int ROW_NUM_PROFIT_AVE = 2;
            const int ROW_NUM_PROFIT_BELOW = 3;

            const int COL_NUM_MARGIN = 6;
            const int ROW_NUM_MARGIN_CLEAN = 1;
            const int ROW_NUM_MARGIN_AVE = 2;
            const int ROW_NUM_MARGIN_BELOW = 3;

            const int COL_LAST_COL_NUM = 6;
            const int ROW_LAST_ROW_NUM = 4;
            
            const int ROWS_TO_SKIP = 4; 

            Excel.Application xl = null;
            Excel._Workbook wb = null;
            Excel._Worksheet sheet = null;
            bool SaveChanges = false;
            Excel.Range chartRange;

            try
            {
                GC.Collect();

                // Create a new instance of Excel from scratch
                xl = new Excel.Application();
                xl.Visible = false;

                // Add one workbook to the instance of Excel

                wb = (Excel._Workbook)(xl.Workbooks.Add(Missing.Value));

                // make sure there are exactly 3 worksheets
                while (wb.Sheets.Count < 3)
                    wb.Sheets.Add();

                while (wb.Sheets.Count > 3)
                    ((Excel._Worksheet)(wb.Sheets[1])).Delete();


                //**********************************************
                // now set up unconditional recipients worksheet
                //**********************************************

                sheet = (Excel._Worksheet)(wb.Sheets[1]);
                sheet.Select();
                sheet.Name = "Car Values";

                //sheet.get_Range("A1").Value = "VRM";
                //sheet.Cells[4, 4] = "test";

                int currRowCounter = 1;
                int numRecsProcessed = 0;

                                
                foreach (var currCar in CarValuations)
                {
                    // add page break after every 7 cars
                    if (currRowCounter > 1 && numRecsProcessed % 7 == 0)
                    {
                        sheet.HPageBreaks.Add(sheet.get_Range("A" + currRowCounter.ToString()));
                    } // end if

                    sheet.Cells[currRowCounter + ROW_NUM_LOT_NO, COL_NUM_LOT_NO] = currCar.Lot_Number;

                    // merge lot number cells
                    sheet.Range[sheet.Cells[currRowCounter + ROW_NUM_LOT_NO, COL_NUM_LOT_NO],
                        sheet.Cells[currRowCounter + ROW_NUM_LOT_NO + ROWS_TO_SKIP - 1, COL_NUM_LOT_NO]].Merge();

                    // horizontally align lot number cells
                    sheet.Range[sheet.Cells[currRowCounter + ROW_NUM_LOT_NO, COL_NUM_LOT_NO],
                        sheet.Cells[currRowCounter + ROW_NUM_LOT_NO + ROWS_TO_SKIP - 1, COL_NUM_LOT_NO]].HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;

                    // vertically align lot number cells
                    sheet.Range[sheet.Cells[currRowCounter + ROW_NUM_LOT_NO, COL_NUM_LOT_NO],
                        sheet.Cells[currRowCounter + ROW_NUM_LOT_NO + ROWS_TO_SKIP - 1, COL_NUM_LOT_NO]].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                    sheet.Cells[currRowCounter + ROW_NUM_DESC, COL_NUM_DESC] = currCar.Long_Description;
                    
                    // merge description cells
                    sheet.Range[sheet.Cells[currRowCounter + ROW_NUM_DESC, COL_NUM_DESC],
                        sheet.Cells[currRowCounter + ROW_NUM_DESC, COL_LAST_COL_NUM]].Merge();

                    // vertically align description cells
                    sheet.Range[sheet.Cells[currRowCounter + ROW_NUM_DESC, COL_NUM_DESC],
                        sheet.Cells[currRowCounter + ROW_NUM_DESC, COL_LAST_COL_NUM]].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                    // wrap description cell text
                    sheet.Range[sheet.Cells[currRowCounter + ROW_NUM_DESC, COL_NUM_DESC],
                        sheet.Cells[currRowCounter + ROW_NUM_DESC, COL_LAST_COL_NUM]].WrapText = true;

                    // add reg no 
                    sheet.Cells[currRowCounter + ROW_NUM_REGNO, COL_NUM_REGNO] = currCar.RegNo;

                    // add valuations
                    sheet.Cells[currRowCounter + ROW_NUM_VALUE_CLEAN, COL_NUM_VALUATIONS] = currCar.CarValuation_Cap.LivePrice_Clean;
                    sheet.Cells[currRowCounter + ROW_NUM_VALUE_AVE, COL_NUM_VALUATIONS] = currCar.CarValuation_Cap.LivePrice_Ave;
                    sheet.Cells[currRowCounter + ROW_NUM_VALUE_BELOW, COL_NUM_VALUATIONS] = currCar.CarValuation_Cap.LivePrice_Below;

                    // add retail valuation
                    sheet.Cells[currRowCounter + ROW_NUM_VALUE_RETAIL, COL_NUM_VALUE_RETAIL] = currCar.CarValuation_Cap.LivePrice_Retail;

                    // merge retail column
                    sheet.Range[sheet.Cells[currRowCounter + ROW_NUM_VALUE_RETAIL, COL_NUM_VALUE_RETAIL],
                        sheet.Cells[currRowCounter + ROW_NUM_VALUE_RETAIL + ROWS_TO_SKIP - 2, COL_NUM_VALUE_RETAIL]].Merge();

                    // vertical align the retail column
                    sheet.Range[sheet.Cells[currRowCounter + ROW_NUM_VALUE_RETAIL, COL_NUM_VALUE_RETAIL],
                        sheet.Cells[currRowCounter + ROW_NUM_VALUE_RETAIL + ROWS_TO_SKIP - 2, COL_NUM_VALUE_RETAIL]].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                    // add profit amount column
                    sheet.Cells[currRowCounter + ROW_NUM_PROFIT_CLEAN, COL_NUM_PROFIT] = string.Format("=${0}{1}-{2}{3}", GetExcelColumnName(COL_NUM_VALUE_RETAIL), (currRowCounter + ROW_NUM_VALUE_RETAIL).ToString(), GetExcelColumnName(COL_NUM_VALUATIONS), (currRowCounter + ROW_NUM_VALUE_CLEAN).ToString());
                    sheet.Cells[currRowCounter + ROW_NUM_PROFIT_AVE, COL_NUM_PROFIT] = string.Format("=${0}{1}-{2}{3}", GetExcelColumnName(COL_NUM_VALUE_RETAIL), (currRowCounter + ROW_NUM_VALUE_RETAIL).ToString(), GetExcelColumnName(COL_NUM_VALUATIONS), (currRowCounter + ROW_NUM_VALUE_AVE).ToString());
                    sheet.Cells[currRowCounter + ROW_NUM_PROFIT_BELOW, COL_NUM_PROFIT] = string.Format("=${0}{1}-{2}{3}", GetExcelColumnName(COL_NUM_VALUE_RETAIL), (currRowCounter + ROW_NUM_VALUE_RETAIL).ToString(), GetExcelColumnName(COL_NUM_VALUATIONS), (currRowCounter + ROW_NUM_VALUE_BELOW).ToString());

                    // add profit margin column
                    sheet.Cells[currRowCounter + ROW_NUM_MARGIN_CLEAN, COL_NUM_MARGIN] = String.Format("=({0}{1}/{2}{3})", GetExcelColumnName(COL_NUM_PROFIT), (currRowCounter + ROW_NUM_PROFIT_CLEAN).ToString(), GetExcelColumnName(COL_NUM_VALUATIONS), (currRowCounter + ROW_NUM_PROFIT_CLEAN).ToString());
                    sheet.Cells[currRowCounter + ROW_NUM_MARGIN_AVE, COL_NUM_MARGIN] = String.Format("=({0}{1}/{2}{3})", GetExcelColumnName(COL_NUM_PROFIT), (currRowCounter + ROW_NUM_PROFIT_AVE).ToString(), GetExcelColumnName(COL_NUM_VALUATIONS), (currRowCounter + ROW_NUM_PROFIT_AVE).ToString());
                    sheet.Cells[currRowCounter + ROW_NUM_MARGIN_BELOW, COL_NUM_MARGIN] = String.Format("=({0}{1}/{2}{3})", GetExcelColumnName(COL_NUM_PROFIT), (currRowCounter + ROW_NUM_PROFIT_BELOW).ToString(), GetExcelColumnName(COL_NUM_VALUATIONS), (currRowCounter + ROW_NUM_PROFIT_BELOW).ToString());
                    
                    // add grid lines
                    chartRange = sheet.get_Range(GetExcelColumnName(COL_FIRST_COL_NUM) + (currRowCounter + ROW_FIRST_ROW_NUM).ToString(), GetExcelColumnName(COL_LAST_COL_NUM) + (currRowCounter + ROW_LAST_ROW_NUM - 1).ToString());
                    chartRange.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    chartRange.Cells.Borders.Color = ColorTranslator.ToOle(Color.Black);
                    chartRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
                    chartRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;
                    chartRange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
                    chartRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                    
                    // increase height of description row
                    sheet.get_Range("A" + (currRowCounter + ROW_FIRST_ROW_NUM).ToString()).RowHeight = 50;

                    // increase width of reg no column
                    sheet.get_Range(GetExcelColumnName(COL_NUM_REGNO) + (currRowCounter + ROW_FIRST_ROW_NUM).ToString()).ColumnWidth = 15;

                    // increase width of valuations column
                    sheet.get_Range(GetExcelColumnName(COL_NUM_VALUATIONS) + (currRowCounter + ROW_FIRST_ROW_NUM).ToString()).ColumnWidth = 15;

                    // set format of valuation columns to currency
                    sheet.get_Range(GetExcelColumnName(COL_NUM_VALUATIONS) + (currRowCounter + ROW_NUM_VALUE_CLEAN).ToString()).NumberFormat = "£#,###";
                    sheet.get_Range(GetExcelColumnName(COL_NUM_VALUATIONS) + (currRowCounter + ROW_NUM_VALUE_AVE).ToString()).NumberFormat = "£#,###";
                    sheet.get_Range(GetExcelColumnName(COL_NUM_VALUATIONS) + (currRowCounter + ROW_NUM_VALUE_BELOW).ToString()).NumberFormat = "£#,###";

                    sheet.get_Range(GetExcelColumnName(COL_NUM_VALUE_RETAIL) + (currRowCounter + ROW_NUM_VALUE_RETAIL).ToString()).NumberFormat = "£#,###";

                    sheet.get_Range(GetExcelColumnName(COL_NUM_PROFIT) + (currRowCounter + ROW_NUM_PROFIT_CLEAN).ToString()).NumberFormat = "£#,###";
                    sheet.get_Range(GetExcelColumnName(COL_NUM_PROFIT) + (currRowCounter + ROW_NUM_PROFIT_AVE).ToString()).NumberFormat = "£#,###";
                    sheet.get_Range(GetExcelColumnName(COL_NUM_PROFIT) + (currRowCounter + ROW_NUM_PROFIT_BELOW).ToString()).NumberFormat = "£#,###";

                    sheet.get_Range(GetExcelColumnName(COL_NUM_MARGIN) + (currRowCounter + ROW_NUM_MARGIN_CLEAN).ToString()).NumberFormat = "###%";
                    sheet.get_Range(GetExcelColumnName(COL_NUM_MARGIN) + (currRowCounter + ROW_NUM_MARGIN_AVE).ToString()).NumberFormat = "###%";
                    sheet.get_Range(GetExcelColumnName(COL_NUM_MARGIN) + (currRowCounter + ROW_NUM_MARGIN_BELOW).ToString()).NumberFormat = "###%";

                    // increase width of retail valuations column
                    sheet.get_Range(GetExcelColumnName(COL_NUM_VALUE_RETAIL) + (currRowCounter + ROW_FIRST_ROW_NUM).ToString()).ColumnWidth = 15;

                    // increase width of profit calcs column
                    sheet.get_Range(GetExcelColumnName(COL_NUM_PROFIT) + (currRowCounter + ROW_FIRST_ROW_NUM).ToString()).ColumnWidth = 15;

                    currRowCounter += ROWS_TO_SKIP; // skip X rows down to take into account merged lines

                    numRecsProcessed++;

                } // end foreach


                //**************************************
                // Let loose control of the Excel instance
                //**************************************

                // ensure the 1st worksheet is selected
                ((Excel._Worksheet)wb.Sheets[1]).Select();

                xl.Visible = false;
                xl.UserControl = false;

                // Set a flag saying that all is well and it is ok to save our changes to a file.
                SaveChanges = true;

                //  Save the file to disk
                wb.SaveAs(fullFileName_ValuationsReport, Excel.XlFileFormat.xlWorkbookNormal,
                          null, null, false, false, Excel.XlSaveAsAccessMode.xlExclusive,
                          false, false, null, null, null);

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message, ex);
            }
            finally
            {

                try
                {
                    // Repeat xl.Visible and xl.UserControl releases just to be sure
                    // we didn't error out ahead of time.

                    if (xl != null)
                    {
                        xl.Visible = false;
                        xl.UserControl = false;

                        if (wb != null)
                        {
                            // Close the document and avoid user prompts to save if our method failed.
                            wb.Close(SaveChanges, null, null);
                        }

                        xl.Workbooks.Close();
                    }
                }
                catch { }

                // Gracefully exit out and destroy all COM objects to avoid hanging instances
                // of Excel.exe whether our method failed or not.

                if (xl != null) { xl.Quit(); }

                if (sheet != null) { Marshal.ReleaseComObject(sheet); }
                if (wb != null) { Marshal.ReleaseComObject(wb); }
                if (xl != null) { Marshal.ReleaseComObject(xl); }

                sheet = null;
                wb = null;
                xl = null;
                GC.Collect();
            } // end try/catch/finally

        }


        /// <summary>
        /// Converts a number into an Excel column name
        /// ie 27 => AA
        /// </summary>
        /// <param name="columnNumber"></param>
        /// <returns></returns>
        public static string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            } // end while

            return columnName;
        }


    }
}
