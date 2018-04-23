using System;
using System.Collections.Generic;

namespace XLExportExample {
    class SalesData 
    {
        public SalesData(string state, double actualSales, double targetSales, double profit, double marketShare) {
            State = state;
            ActualSales = actualSales;
            TargetSales = targetSales;
            Profit = profit;
            MarketShare = marketShare;
        }

        public string State { get; private set; }
        public double ActualSales { get; private set; }
        public double TargetSales { get; private set; }
        public double Profit { get; private set; }
        public double MarketShare { get; private set; }
    }

    static class SalesRepository {
        static string[] states = new string[] { 
            "Alabama", "Arizona", "California", "Colorado", "Connecticut", "Florida", "Georgia", "Idaho", 
            "Illinois", "Indiana", "Kentucky", "Maine", "Massachusetts", "Michigan", "Minnesota", "Mississippi", 
            "Missouri", "Montana", "Nevada", "New Hampshire", "New Mexico", "New York", "North Carolina", "Ohio", 
            "Oregon", "Rhode Island", "South Carolina", "South Dakota", "Tennessee", "Texas", "Utah", "Virginia", 
            "Washington", "Wisconsin", "Wyoming"};

        public static List<SalesData> GetSalesData() {
            Random random = new Random();
            List<SalesData> result = new List<SalesData>();
            foreach(string state in states) {
                double targetSales = (random.NextDouble() * 500 + 40) * 1e6;
                double actualSales = targetSales * (0.9 + random.NextDouble() * 0.2);
                double profit = actualSales * (random.NextDouble() * 0.1 - 0.03);
                if (Math.Abs(profit) < 1e6)
                    profit = Math.Sign(profit) * 1e6;
                double marketShare = random.NextDouble() * 0.2 + 0.1;
                result.Add(new SalesData(state, actualSales, targetSales, profit, marketShare));
            }
            return result;
        }
    }
}
