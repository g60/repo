using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.Text.RegularExpressions;
using System.IO;

/*
 * 
 * This takes a web page valuation from Honest John and returns the prices from it
 * 
 * */

namespace Car_Price_Guider
{
    class ValuationParser_HonestJohn
    {

        public static void RequestValuationFromWebsite(CarDetails carDetails, string Filename)
        {
            ValuationParser_HonestJohn valuer = new ValuationParser_HonestJohn();

            string baseURL = @"http://www.honestjohn.co.uk/used-prices/";
            string args = "?q=";
            string detailsToFind = "Ford Mondeo 2004 Automatic Estate Ghia X TDCi";
            detailsToFind = "renault Clio 2003 extreme";

            detailsToFind = carDetails.FormatForValuation_HonestJohn();

            string fullURL = baseURL + args + detailsToFind;

            var req = (HttpWebRequest)WebRequest.Create(fullURL);
            req.Method = "GET";

            using (WebResponse odpoved = req.GetResponse())
            {
                
                //Console.WriteLine(DateTime.Now.ToShortDateString());
                var fileStream = File.Create(Filename);

                Stream str = odpoved.GetResponseStream();
                str.CopyTo(fileStream);
                fileStream.Close();
            }

            return;



        }

        public static CarValuation_HonestJohn GetValuation_StoredFile(string FileName)
        {

            ValuationParser_HonestJohn valuer = new ValuationParser_HonestJohn();

            CarValuation_HonestJohn returnPrices = new CarValuation_HonestJohn();

            var req1 = (FileWebRequest)WebRequest.Create(FileName);
            req1.Method = "GET";

            using (WebResponse odpoved = req1.GetResponse())
            {
                var htmlDoc = new HtmlAgilityPack.HtmlDocument();
                htmlDoc.Load(odpoved.GetResponseStream());

                // All this below will format the text in 1 block
                string allPriceText = valuer.GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]").Trim(); // this works atm

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

                }

                var result = Regex.Split(allPriceText, "\r\n|\r|\n");

                /*
                 * 
                 * At a dealer, aim to pay between £7,350 and £9,300 
                    At a franchised dealer, expect to pay £8,550. 
                    At an independent dealer, expect to pay £8,000. 
                    At a car supermarket, expect to pay £7,900. 
                    From a private seller, aim to pay between  £6,950 and £8,150 and expect to pay £7,500. 
                    If selling in part exchange, expect to receive £6,530.
                 * 
                 * 
                */


                //CarValuation returnPrices = new CarValuation();


                for (int i = 0; i < result.Length; i++)
                {
                    string currLine = result[i];
                    currLine = currLine.Replace("At an", "");
                    currLine = currLine.Replace("At a", "");
                    currLine = currLine.Replace("From a", "");
                    currLine = currLine.Replace("If selling in", "");
                    Match regex_Res_SellerType = Regex.Match(currLine, @"^[a-zA-Z ]*");
                    string sellerType = regex_Res_SellerType.Value;

                    if (sellerType != "")
                    {
                        MatchCollection prices = Regex.Matches(currLine, @"£[0-9,]*");

                        switch (sellerType.ToUpper().Trim())
                        {
                            case "DEALER":
                                returnPrices.DealerMinPrice = prices[0].Value;
                                returnPrices.DealerMaxPrice = prices[1].Value;
                                if (prices.Count == 3)
                                {
                                    returnPrices.DealerExpectedPrice = prices[2].Value;
                                }
                                break;
                            case "FRANCHISED DEALER":
                                returnPrices.DealerFranchisedExpectedPrice = prices[0].Value;
                                break;
                            case "INDEPENDENT DEALER":
                                returnPrices.DealerIndependantExpectedPrice = prices[0].Value;
                                break;
                            case "CAR SUPERMARKET":
                                returnPrices.CarSupermarketExpectedPrice = prices[0].Value;
                                break;
                            case "PRIVATE SELLER":
                                returnPrices.PrivateSellerMinPrice = prices[0].Value;
                                returnPrices.PrivateSellerMaxPrice = prices[1].Value;
                                if (prices.Count == 3)
                                {
                                    returnPrices.PrivateSellerExpectedPrice = prices[2].Value;
                                }
                                break;
                            case "PART EXCHANGE":
                                returnPrices.PartExPrice = prices[0].Value;
                                break;
                            default:
                                break;
                        } // end switch
                    } // end if

                } // end for

            } // end using

            return returnPrices;

        }

        /// <summary>
        /// Old method, uses xpath instead of regex
        /// </summary>
        /// <param name="FileName"></param>
        /// <returns></returns>
        public static CarValuation_HonestJohn GetValuation_StoredFile1(string FileName)
        {

            ValuationParser_HonestJohn valuer = new ValuationParser_HonestJohn();

            CarValuation_HonestJohn returnPrices = new CarValuation_HonestJohn();

            var req1 = (FileWebRequest)WebRequest.Create(FileName);
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

                string aveAdvertPrice_Min = valuer.GetPathText(htmlDoc, "//*[@id=\"prices\"]/div[3]/div/div/div[3]/p[1]/b[1]");
                string aveAdvertPrice_Max = valuer.GetPathText(htmlDoc, "//*[@id=\"prices\"]/div[3]/div/div/div[3]/p[1]/b[2]");

                string typeName1 = valuer.GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]/tr[1]/td[1]/b[1]");
                string typeName2 = valuer.GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]/tr[2]/td[1]/b[1]");
                string typeName3 = valuer.GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]/tr[3]/td[1]/b[1]");
                string typeName4 = valuer.GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]/tr[4]/td[1]/b[1]");
                string typeName5 = valuer.GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]/tr[5]/td[1]/b[1]");
                string typeName6 = valuer.GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]/tr[6]/td[1]/b[1]");


                string type1MinPrice = valuer.GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]/tr[1]/td[3]/b[1]");
                string type1MaxPrice = valuer.GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]/tr[1]/td[3]/b[2]");
                string type1ExpPrice = valuer.GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]/tr[2]/td[3]/b[1]");

                string type2MinPrice = valuer.GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]/tr[2]/td[3]/b[1]");
                string type2MaxPrice = valuer.GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]/tr[2]/td[3]/b[2]");
                string type2ExpPrice = valuer.GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]/tr[2]/td[3]/b[1]");

                string type3MinPrice = valuer.GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]/tr[3]/td[3]/b[1]");
                string type3MaxPrice = valuer.GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]/tr[3]/td[3]/b[2]");
                string type3ExpPrice = valuer.GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]/tr[4]/td[2]/b[1]");
                string type3ExpPrice1 = valuer.GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]/tr[3]/td[3]/b[1]");

                string type4MinPrice = valuer.GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]/tr[4]/td[3]/b[1]");
                string type4MaxPrice = valuer.GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]/tr[4]/td[3]/b[2]");
                string type4ExpPrice = valuer.GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]/tr[5]/td[2]/b[1]");

                string type5ExpectedPrice = valuer.GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]/tr[5]/td[3]/b[1]");

                string type6ExpectedPrice = valuer.GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]/tr[6]/td[3]/b[1]");

                string partExPrice = valuer.GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]/tr[6]/td[3]/b[1]");

                Console.WriteLine("#####################################");
                Console.WriteLine("#####################################");

                Console.WriteLine("Most adverts are between " + aveAdvertPrice_Min + " and " + aveAdvertPrice_Max);

                Console.WriteLine("Type1: " + typeName1);
                Console.WriteLine("Type2: " + typeName2);
                Console.WriteLine("Type3: " + typeName3);
                Console.WriteLine("Type4: " + typeName4);
                Console.WriteLine("Type5: " + typeName5);
                Console.WriteLine("Type6: " + typeName6);

                Console.WriteLine("type1MinPrice: " + type1MinPrice);
                Console.WriteLine("type1MaxPrice: " + type1MaxPrice);
                Console.WriteLine("type1ExpPrice: " + type1ExpPrice);

                Console.WriteLine("type2MinPrice: " + type2MinPrice);
                Console.WriteLine("type2MaxPrice: " + type2MaxPrice);
                Console.WriteLine("type2ExpPrice: " + type2ExpPrice);

                Console.WriteLine("type3MinPrice: " + type3MinPrice);
                Console.WriteLine("type3MaxPrice: " + type3MaxPrice);
                Console.WriteLine("type3ExpPrice: " + type3ExpPrice);
                Console.WriteLine("type3ExpPrice1: " + type3ExpPrice1);

                Console.WriteLine("type4MinPrice: " + type4MinPrice);
                Console.WriteLine("type4MaxPrice: " + type4MaxPrice);
                Console.WriteLine("type4ExpPrice: " + type4ExpPrice);

                Console.WriteLine("type5ExpectedPrice: " + type5ExpectedPrice);

                Console.WriteLine("type6ExpectedPrice: " + type6ExpectedPrice);

                Console.WriteLine("#####################################");
                Console.WriteLine("#####################################");

                if (typeName1 == "dealer")
                {
                    returnPrices.DealerMinPrice = type1MinPrice;
                    returnPrices.DealerMaxPrice = type1MaxPrice;

                    if (typeName2 == null || typeName2 == "")
                    {
                        returnPrices.DealerExpectedPrice = type2ExpPrice;
                    } // end if

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
                else if (typeName2 == "independent dealer")
                {
                    returnPrices.DealerIndependantExpectedPrice = type2ExpPrice;
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
                else if (typeName3 == "independent dealer")
                {
                    returnPrices.DealerIndependantExpectedPrice = type3ExpPrice1;
                }
                else if (typeName3 == "private seller")
                {
                    returnPrices.PrivateSellerMinPrice = type3MinPrice;
                    returnPrices.PrivateSellerMaxPrice = type3MaxPrice;
                    returnPrices.PrivateSellerExpectedPrice = type3ExpPrice;
                }


                if (typeName4 == "dealer")
                {
                    returnPrices.DealerMinPrice = type4MinPrice;
                    returnPrices.DealerMaxPrice = type4MaxPrice;
                }
                else if (typeName4 == "franchised dealer")
                {
                    //returnPrices.DealerFranchisedExpectedPrice = type4ExpPrice;
                }
                else if (typeName4 == "private seller")
                {
                    returnPrices.PrivateSellerMinPrice = type4MinPrice;
                    returnPrices.PrivateSellerMaxPrice = type4MaxPrice;
                    returnPrices.PrivateSellerExpectedPrice = type4ExpPrice;
                }


                if (typeName5 == "part exchange")
                {
                    returnPrices.PartExPrice = type5ExpectedPrice;
                }

                if (typeName6 == "part exchange")
                {
                    returnPrices.PartExPrice = type6ExpectedPrice;
                }

                //returnPrices.DealerFranchisedExpectedPrice = dealerFranchisedExpectedPrice;
                //returnPrices.DealerIndependantExpectedPrice = dealerIndependantExpectedPrice;



                //returnPrices.PrivateSellerExpectedPrice = privateSellerExpectedPrice;

                //returnPrices.PartExPrice = partExPrice;



                // All this below will format the text in 1 block
                string allPriceText = valuer.GetPathText(htmlDoc, "//*[@class=\"pricePoints\"]").Trim(); // this works atm

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
