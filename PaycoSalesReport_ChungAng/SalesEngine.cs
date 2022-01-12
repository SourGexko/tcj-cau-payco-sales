using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PaycoSalesReport_ChungAng
{
    public class SalesEngine
    {
        const string SEOUL_CAUBURGER_KEY = "EDCABT";
        const string SEOUL_RESTAURANT_KEY = "XBMM3X";
        const string ANSUNG_CAUBURGER_KEY = "PLBADQ";
        const string ANSUNG_RESTAURANT_KEY = "VTE1B9";
        public List<Deal> SeoulCauburgerDeals { get; set; } = new List<Deal>();
        public List<Deal> SeoulRestaurantDeals { get; set; } = new List<Deal>();
        public List<Deal> AnsungCauburgerDeals { get; set; } = new List<Deal>();
        public List<Deal> AnsungRestaurantDeals { get; set; } = new List<Deal>();
        public List<DateSale> SeoulCauburgerDateSales { get; set; } = new List<DateSale>();
        public List<DateSale> SeoulRestaurantDateSales { get; set; } = new List<DateSale>();
        public List<DateSale> AnsungCauburgerDateSales { get; set; } = new List<DateSale>();
        public List<DateSale> AnsungRestaurantDateSales { get; set; } = new List<DateSale>();

        public void Encapsulation(string filePath)
        {
            FileInfo existingFile = new FileInfo(filePath);
            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                var worksheet = package.Workbook.Worksheets[0];
                var rowCount = worksheet.Dimension.Rows;
                for (int i = 2; i <= rowCount; ++i)
                {
                    double date = double.Parse(worksheet.GetValue(i, 1).ToString());
                    DateTime conv = DateTime.FromOADate(date);
                    var name = worksheet.GetValue(i, 2);
                    var paycoPoint = worksheet.GetValue(i, 13);
                    var storeKey = worksheet.GetValue(i, 3);
                    switch (storeKey)
                    {
                        case SEOUL_CAUBURGER_KEY:
                            SeoulCauburgerDeals.Add(new Deal(conv, ELocation.Seoul, name.ToString(), int.Parse(paycoPoint.ToString())));
                            Console.WriteLine(conv.ToString("yyyyMMdd") + " / 서울 카우버거 / " + storeKey + " / " + paycoPoint);
                            break;
                        case SEOUL_RESTAURANT_KEY:
                            SeoulRestaurantDeals.Add(new Deal(conv, ELocation.Seoul, name.ToString(), int.Parse(paycoPoint.ToString())));
                            Console.WriteLine(conv.ToString("yyyyMMdd") + " / 서울 참슬기 / " + storeKey + " / " + paycoPoint);
                            break;
                        case ANSUNG_CAUBURGER_KEY:
                            AnsungCauburgerDeals.Add(new Deal(conv, ELocation.Ansung, name.ToString(), int.Parse(paycoPoint.ToString())));
                            Console.WriteLine(conv.ToString("yyyyMMdd") + " / 안성 카우버거/ " + storeKey + " / " + paycoPoint);
                            break;
                        case ANSUNG_RESTAURANT_KEY:
                            AnsungRestaurantDeals.Add(new Deal(conv, ELocation.Ansung, name.ToString(), int.Parse(paycoPoint.ToString())));
                            Console.WriteLine(conv.ToString("yyyyMMdd") + " / 안성 참슬기/ " + storeKey + " / " + paycoPoint);
                            break;
                        default:
                            break;
                    }
                }
            }
        }

        public void DevideSalesPerMonthByDate(int year, int month)
        {
            var days = DateTime.DaysInMonth(year, month);
            for (int i = 1; i <= days; ++i)
            {
                var dateTime = new DateTime(year, month, i);
                var seoulCauburgerTotalPricePerDate = SeoulCauburgerDeals.Where(d => d.DateTime.Date == dateTime).Sum(d => d.Price);
                var seoulRestaurantTotalPricePerDate = SeoulRestaurantDeals.Where(d => d.DateTime.Date == dateTime).Sum(d => d.Price);
                var ansungCauburgerTotalPricePerDate = AnsungCauburgerDeals.Where(d => d.DateTime.Date == dateTime).Sum(d => d.Price);
                var ansungRestaurantTotalPricePerDate = AnsungRestaurantDeals.Where(d => d.DateTime.Date == dateTime).Sum(d => d.Price);
                SeoulCauburgerDateSales.Add(new DateSale(ELocation.Seoul, EStoreType.Cauburger, dateTime, seoulCauburgerTotalPricePerDate));
                SeoulRestaurantDateSales.Add(new DateSale(ELocation.Seoul, EStoreType.Restaurant, dateTime, seoulRestaurantTotalPricePerDate));
                AnsungCauburgerDateSales.Add(new DateSale(ELocation.Ansung, EStoreType.Cauburger, dateTime, ansungCauburgerTotalPricePerDate));
                AnsungRestaurantDateSales.Add(new DateSale(ELocation.Ansung, EStoreType.Restaurant, dateTime, ansungRestaurantTotalPricePerDate));
            }
        }

        public void WriteToExcel(int year, int month)
        {
            string folderPath = @"output";
            DirectoryInfo di = new DirectoryInfo(folderPath);

            if (!di.Exists)
            {
                di.Create();
            }


            List<DateSale> totalDateSales = SeoulCauburgerDateSales.Concat(SeoulRestaurantDateSales).Concat(AnsungCauburgerDateSales).Concat(AnsungRestaurantDateSales).ToList();
            FileInfo template = new FileInfo("Template.xlsx");
            using (var package = new ExcelPackage(template))
            {
                var worksheet = package.Workbook.Worksheets[0];
                for (int i = 0; i < totalDateSales.Count; ++i)
                {
                    string storeType = String.Empty;
                    switch (totalDateSales[i].StoreType)
                    {
                        case EStoreType.Cauburger:
                            storeType = "카우버거";
                            break;
                        case EStoreType.Restaurant:
                            storeType = "참슬기";
                            break;
                        default:
                            break;
                    }
                    worksheet.Cells[i + 2, 1].Value = totalDateSales[i].Location.ToString();
                    worksheet.Cells[i + 2, 2].Value = storeType;
                    worksheet.Cells[i + 2, 3].Value = totalDateSales[i].DateTime.ToString("yyyy년 MM월 dd일");
                    worksheet.Cells[i + 2, 4].Value = totalDateSales[i].Price;
                }
                var newFile = new FileInfo(@$"{year}{month}_cau_payco_sales.xlsx");
                package.SaveAs(newFile);
            }
        }
    }
}
