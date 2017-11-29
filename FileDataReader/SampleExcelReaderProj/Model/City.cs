using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SampleExcelReaderProj.Model
{
    public class City
    {
        public string  CityName { get; set; }

        public string  StateCode { get; set; }

        public string  CountryCode { get; set; }

        public string Latitude { get; set; }

        public string Longitude { get; set; }

        public string IsEnabled { get; set; }

        public string IataCityCode { get; set; }

        public string FullTextColumn { get; set; }

    }
}
