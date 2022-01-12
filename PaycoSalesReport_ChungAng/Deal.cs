using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PaycoSalesReport_ChungAng
{
    public class Deal
    {
        public DateTime DateTime { get; private set; }
        public ELocation Location { get; private set; }
        public string StoreName { get; set; }
        public int Price { get; set; }

        public Deal(DateTime dateTime, ELocation location, string storeName, int price)
        {
            DateTime = dateTime;
            Location = location;
            StoreName = storeName;
            Price = price;
        }
    }
}
