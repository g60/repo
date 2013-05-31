using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using HtmlAgilityPack;

namespace Car_Price_Guider
{
    class AuctionCatalogueParser_Bawtry
    {

        public static List<CarDetails> ParseCatalogue_StoredFile(string FileName)
        {

            AuctionCatalogueParser_Bawtry parser = new AuctionCatalogueParser_Bawtry();

            List<CarDetails> returnDetails = new List<CarDetails>();

            var req1 = (FileWebRequest)WebRequest.Create(FileName);
            req1.Method = "GET";

            using (WebResponse odpoved = req1.GetResponse())
            {
                var htmlDoc = new HtmlAgilityPack.HtmlDocument();
                htmlDoc.Load(odpoved.GetResponseStream());

                //string test = valuer.GetPathText(htmlDoc, "//*[@id=\"content\"]/table/tbody/tr[999]/td[2]");

                for (int i = 2; i < 999; i+=2)
                {

                    string carDetail = parser.GetPathText(htmlDoc, "//*[@id=\"content\"]/table/tbody/tr[" + i + "]/td[2]");

                    if (carDetail != null)
                    {

                        var result = Regex.Split(carDetail, "\r\n|\r|\n");

                        string longDesc1 = null;
                        string longDesc2 = null;
                        string registered_raw = null;

                        
                        for (int j = 0; j < result.Length; j++)
                        {
                            string currLine = result[j].Trim();
                            if (currLine != "")
                            {
                                //Console.WriteLine(currLine);
                                if (longDesc1 == null)
                                {
                                    longDesc1 = currLine;
                                }
                                else if (longDesc2 == null)
                                {
                                    longDesc2 = currLine;
                                }
                                else if (registered_raw == null)
                                {
                                    registered_raw = currLine;
                                }
                                else
                                {
                                    throw new Exception("Unexpected detail line");
                                }
                            }
                        }

                        CarDetails newCar = new CarDetails();
                        newCar.Long_Description = longDesc1 + " " + longDesc2;

                        registered_raw = registered_raw.Replace("Registered", "");
                        registered_raw = registered_raw.Trim();

                        DateTime RegDate;

                        if (DateTime.TryParse(registered_raw, out RegDate))
                        {
                            newCar.RegDate = RegDate;
                        }

                        string lotNo = parser.GetPathText(htmlDoc, "//*[@id=\"content\"]/table/tbody/tr[" + i + "]/td[4]");
                        
                        newCar.Lot_Number = lotNo;


                        //*[@id=\"content\"]/table/tbody/tr[4]/td/div
                        //*[@id="content"]/table/tbody/tr[4]/td/div/img[1]
                        //*[@id="content"]/table/tbody/tr[4]/td/div/img[2]

                        string regNo = "";

                        for (int x = 1; x < 15; x++)
                        {
                            string xPath = "//*[@id=\"content\"]/table/tbody/tr[" + (i + 1) +"]/td/div/img[" + x + "]";

                            var nodes = htmlDoc.DocumentNode.SelectNodes(xPath);

                            if (nodes != null)
                            {
                                for (int y = 0; y < nodes.Count; y++)
                                {
                                    regNo = regNo + nodes[y].Attributes[1].Value;
                                } // end for
                            } // end if

                        } // end for

                        //Console.WriteLine();

                        newCar.RegNo = regNo.ToUpper();

                        //var nodes = htmlDoc.DocumentNode.SelectNodes(xPath);
                        ////src = new List<string>(nodes.Count);

                        //if (nodes != null)
                        //{
                        //    foreach (var node in nodes)
                        //    {
                        //        if (node.Id != null)
                        //        {
                        //            //src.Add(node.Id);
                        //            //Console.WriteLine("Most adverts are between " + node.InnerText);
                        //            string textFound = node.InnerText.Trim();
                        //            textFound = textFound.Replace("Â", "");
                        //            return textFound;
                        //        }

                        //    }
                        //}
                        

                        //Console.WriteLine("***********************************");
                        //Console.WriteLine("Desc: " + newCar.Long_Description);
                        //Console.WriteLine("Reg date: " + newCar.RegDate.ToShortDateString());
                        //Console.WriteLine("Reg year: " + newCar.RegDate.Year);
                        //Console.WriteLine("Formatted: " + newCar.FormatForValuation_HonestJohn());
                        //Console.WriteLine("***********************************");

                        //Console.WriteLine(carDetail.Trim());

                        returnDetails.Add(newCar);

                    } // end if

                }

                /*
                foreach(HtmlNode link in htmlDoc.DocumentNode.SelectNodes("//*[@id=\"content\"]/table/tbody/tr"))
                 {
                     Console.WriteLine("***********************************");
                     Console.WriteLine(link.InnerText);
                     //Console.WriteLine(link.ToString());
                    //HtmlAttribute att = link["href"];
                    //att.Value = FixLink(att);
                     Console.WriteLine("***********************************");
                 }
                */
            }

            return returnDetails;

        }





        public static List<CarDetails> ParseCatalogue_StoredFile_PATHTESTER(string FileName, string XPath)
        {

            AuctionCatalogueParser_Bawtry parser = new AuctionCatalogueParser_Bawtry();

            List<CarDetails> returnDetails = new List<CarDetails>();

            var req1 = (FileWebRequest)WebRequest.Create(FileName);
            req1.Method = "GET";

            using (WebResponse odpoved = req1.GetResponse())
            {
                var htmlDoc = new HtmlAgilityPack.HtmlDocument();
                htmlDoc.Load(odpoved.GetResponseStream());

                //string test = valuer.GetPathText(htmlDoc, "//*[@id=\"content\"]/table/tbody/tr[999]/td[2]");

                //for (int i = 1; i < 999; i += 2)
                //{

                    //string carDetail = parser.GetPathText(htmlDoc, "//*[@id=\"content\"]/table/tbody/tr[" + i + "]/td[2]");
                    string carDetail = parser.GetPathText(htmlDoc, XPath);

                    if (carDetail != null)
                    {
                        Console.WriteLine("*********************************");
                        Console.WriteLine(carDetail);
                        Console.WriteLine("*********************************");

                    } // end if

                //}

            }

            return returnDetails;

        }






        private string GetPathText(HtmlAgilityPack.HtmlDocument htmlDoc, string xPath)
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


    }
}
