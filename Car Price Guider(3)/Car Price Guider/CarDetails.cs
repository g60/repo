using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Car_Price_Guider
{
    class CarDetails
    {
        public string RegNo { get; set; }
        public DateTime RegDate { get; set; }
        public string Make { get; set; }
        public string Model { get; set; }
        public string Long_Description { get; set; }
        public string Lot_Number { get; set; }

        public string FormatForValuation_HonestJohn()
        {
            return RegDate.Year + " " + Long_Description;
        }

    }
}
