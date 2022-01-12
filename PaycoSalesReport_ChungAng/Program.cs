using System;
using System.IO;
using OfficeOpenXml;

namespace PaycoSalesReport_ChungAng
{
    class Program
    {
        static void Main(string[] args)
        {
            SalesEngine salesEngine = new SalesEngine();
            Console.WriteLine("************* CAU Payco 매출 정산 Program *************");
            Console.WriteLine("* 매달 10일 중앙대 Payco 매출을 전송합니다.");
            Console.WriteLine("* Payco Patners 사이트 (https://master.payco.com/)에서 ");
            Console.WriteLine("* 매출내역 -> 가맹점매출조회 화면에서 전 달 조회 후 조회결과 다운로드 하여");
            Console.Write("* xlsx 파일 경로를 입력하세요: ");
            string filePath = Console.ReadLine();
            try
            {
                salesEngine.Encapsulation(filePath);
            }
            catch (NullReferenceException e)
            {
                Console.WriteLine(e.Message);
            }

            Console.Write("* 년도를 입력하세요: ");
            int year = int.Parse(Console.ReadLine());

            Console.Write("* 월을 입력하세요: ");
            int month = int.Parse(Console.ReadLine());

            Console.WriteLine("* 일자별로 나누는 중...");
            salesEngine.DevideSalesPerMonthByDate(year, month);
            Console.WriteLine("* 완료");

            Console.WriteLine("* 새로운 파일 생성 중...");
            salesEngine.WriteToExcel(year, month);
            Console.WriteLine("* 완료");
            Console.WriteLine("* Program과 동일한 폴더 내에 xlsx파일을 확인해주세요.");

        }

    }
}
