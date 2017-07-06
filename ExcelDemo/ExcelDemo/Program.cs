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
            for (int i = 1; i <= 100; i++)
            {
                var j = i / 2;
                
                var order = new DedicatedOrderExportSourceModel
                {
                    Adcode=$"510100{j}",
                    CarType=$"拼车{j}",
                    CityName=$"成都{j}",
                    CreateTime=DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                    CreatorName=$"测试{j}",
                    DepartmentName=$"测试部门{j}",
                    Destination=$"目的{j}",
                    FlighTime=DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                    FlightNumber=$"3u889{j}",
                    OrderId= j.ToString(),
                    OrderState=$"预约成功{j}",
                    OrderType=$"接机{j}",
                    PeopleNumber=j.ToString(),
                    Price=j.ToString(),
                    ReadState="ReadState",
                    Starting=$"起始地{j}",
                    UseTime=DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
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