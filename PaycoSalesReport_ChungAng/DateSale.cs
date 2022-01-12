using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PaycoSalesReport_ChungAng
{
    public class DateSale
    {
        public ELocation Location { get; private set; }
        public EStoreType StoreType { get; private set; }
        public DateTime DateTime { get; private set; }
        public int Price { get; private set; }

        public DateSale(ELocation location, EStoreType storeType, DateTime dateTime, int price)
        {
            Location = location;
            StoreType = storeType;
            DateTime = dateTime;
            Price = price;
        }
    }
}
