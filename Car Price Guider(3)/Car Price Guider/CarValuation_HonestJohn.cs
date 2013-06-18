using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Car_Price_Guider
{
    class CarValuation_HonestJohn
    {

        private string dealerMinPrice;
        private string dealerMaxPrice;
        private string dealerExpectedPrice;
        private string dealerFranchisedExpectedPrice;
        private string dealerIndependantExpectedPrice;
        private string carSupermarketExpectedPrice;
        private string privateSellerMinPrice;
        private string privateSellerMaxPrice;
        private string privateSellerExpectedPrice;
        private string partExPrice;
        private string allPriceText;

        
        public CarValuation_HonestJohn()
        {

        }

        public string DealerMinPrice
        {
            get { return dealerMinPrice; }
            set { dealerMinPrice = value; }
        }
        
        public string DealerMaxPrice
        {
            get { return dealerMaxPrice; }
            set { dealerMaxPrice = value; }
        }

        public string DealerExpectedPrice
        {
            get { return dealerExpectedPrice; }
            set { dealerExpectedPrice = value; }
        }

        public string DealerFranchisedExpectedPrice
        {
            get { return dealerFranchisedExpectedPrice; }
            set { dealerFranchisedExpectedPrice = value; }
        }
        
        public string DealerIndependantExpectedPrice
        {
            get { return dealerIndependantExpectedPrice; }
            set { dealerIndependantExpectedPrice = value; }
        }

        public string CarSupermarketExpectedPrice
        {
            get { return carSupermarketExpectedPrice; }
            set { carSupermarketExpectedPrice = value; }
        }
        
        public string PrivateSellerMinPrice
        {
            get { return privateSellerMinPrice; }
            set { privateSellerMinPrice = value; }
        }
        
        public string PrivateSellerMaxPrice
        {
            get { return privateSellerMaxPrice; }
            set { privateSellerMaxPrice = value; }
        }
        
        public string PrivateSellerExpectedPrice
        {
            get { return privateSellerExpectedPrice; }
            set { privateSellerExpectedPrice = value; }
        }
        
        public string PartExPrice
        {
            get { return partExPrice; }
            set { partExPrice = value; }
        }

        public string AllPriceText
        {
            get { return allPriceText; }
            set { allPriceText = value; }
        }

    }
}
