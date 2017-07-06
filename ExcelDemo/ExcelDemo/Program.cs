using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http.Headers;
using ExcelDemo.Unitity;
using Npoi.Core.SS.Formula.Functions;

namespace ExcelDemo
{   
    class Program
    {
        static void Main(string[] args)
        {
            var path= Directory.GetCurrentDirectory();
            var configpath = Path.Combine(path, "exportTemplate/StationOrderExportTemplate.xml");
            
            var source = new List<DedicatedOrderExportSourceModel>();
            for (int i = 0; i < 10; i++)
            {
                var order = new DedicatedOrderExportSourceModel
                {
                    Adcode="510100",
                    CarType="拼车",
                    CityName="成都",
                    CreateTime=DateTime.Now,
                    CreatorName="测试",
                    DepartmentName="测试部门",
                    Destination=$"目的{i}",
                    FlighTime=DateTime.Now.ToString("yyyy-MM-dd"),
                    FlightNumber=$"3u889{i}",
                    OrderId= i.ToString(),
                    OrderState="预约成功",
                    OrderType="接机",
                    PeopleNumber=i,
                    Price=i,
                    ReadState="ReadState",
                    Starting=$"起始地{i}",
                    UseTime=DateTime.Now
                };
                source.Add(order);

            }
            var excelHelper = new NopiExcelHelper<DedicatedOrderExportSourceModel>("接送机订单","sheet1");
            excelHelper.ExportToFile(configpath,source, "D:\\excel\\order.xls");
            Console.WriteLine("export success");
            Console.ReadLine();
        }
    }
   
}