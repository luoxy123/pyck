using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelDemo
{
    public class DedicatedOrderExportSourceModel
    {
        public string OrderId { get; set; }

        public string Adcode { get; set; }

        public string CityName { get; set; }

        public DateTime CreateTime { get; set; }

        public string CreatorName { get; set; }

        public string DepartmentName { get; set; }

        public string OrderType { get; set; }


        public string CarType { get; set; }


        public string OrderState { get; set; }


        public int PeopleNumber { get; set; }


        public int Price { get; set; }
        public string ReadState { get; set; }


        public string FlightNumber { get; set; }


        public string Starting { get; set; }


        public string Destination { get; set; }


        public DateTime UseTime { get; set; }

        public string FlighTime { get; set; }
    }
}
