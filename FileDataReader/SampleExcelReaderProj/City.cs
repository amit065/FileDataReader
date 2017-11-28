using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SampleExcelReaderProj
{
    public class City
    {
        public string  CityName { get; set; }

        public string  StateCode { get; set; }

        public string  CountryCode { get; set; }

        public float Latitude { get; set; }

        public float Longitude { get; set; }

        public bool IsEnabled { get; set; }

        public string IataCityCode { get; set; }

        public string FullTextColumn { get; set; }

        public string TimeZoneMAppingId { get; set; }
    }
}
