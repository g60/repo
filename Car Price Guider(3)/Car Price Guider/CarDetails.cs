using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace Car_Price_Guider
{
    class CarDetails
    {
        private string _mileage;
        private string _longDescription;

        public string RegNo { get; set; }
        public string Type { get; set; }
        public DateTime RegDate { get; set; }
        public string Make { get; set; }
        public string Model { get; set; }
        public string Fuel { get; set; }
        public string Trans { get; set; }
        public string Doors { get; set; }
        //public string Long_Description { get; set; }
        public string Lot_Number { get; set; }

        public string FromCatalogue { get; set; }

        public CarValuation_CAP CarValuation_Cap { get; set; }
        public CarValuation_HonestJohn CarValuation_HonestJohn { get; set; }

        public string Long_Description
        {
            get
            {
                switch (FromCatalogue.ToUpper())
                {
                    case "BAWTRY":
                        return _longDescription;
                        
                    case "NEWARK":
                        return Make + " " + Model + " " + Fuel + " " + Trans + " " + Doors + " door " + _mileage + " miles";
                        
                    default:
                        return "";
                };
            }
            set
            {
                switch (FromCatalogue.ToUpper())
                {
                    case "BAWTRY":
                        _longDescription = value;
                        break;
                    case "NEWARK":
                        _longDescription = value;
                        break;
                    default:
                        break;
                }
            }
        }

        public string Mileage
        {
            get
            {
                return _mileage;
            }

            set
            {
                switch (FromCatalogue.ToUpper())
                {
                    case "BAWTRY":
                        _mileage = "";
                        _mileage = Regex.Match(value, @"[0-9,]* miles").Value;
                        _mileage = _mileage.Replace("miles", "");
                        _mileage = _mileage.Replace(" ", ""); // remove spaces
                        _mileage = _mileage.Replace(",", ""); // remove commas
                        break;
                    case "NEWARK":
                        _mileage = value;
                        _mileage.Trim();
                        _mileage = _mileage.Replace(" ", ""); // remove spaces
                        _mileage = _mileage.Replace(",", ""); // remove commas
                        break;
                    default:
                        break;
                }

            }
        }

        public string FormatForValuation_HonestJohn()
        {
            return RegDate.Year + " " + Long_Description;
        }

        public string GetMileage()
        {
            if (String.Compare(FromCatalogue, "BAWTRY", true) == 0)
            {
                string mileage = Regex.Match(Long_Description, @"[0-9,]* miles").Value;
                mileage = mileage.Replace("miles", "");
                mileage = mileage.Replace(" ", "");
                mileage = mileage.Replace(",", "");

                return mileage;
            }
            else
            {
                return "";
            } // end if-then-else
        }

        public string FormatForValuation_CAP_Email()
        {
            return "value " + RegNo + " " + GetMileage();
        }

    }
}
