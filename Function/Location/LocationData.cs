using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProjectExcelReader.Function
{
    public class LocationData
    {
        public string locationName { get; set; }
        public string countryName { get; set; }
        public string image { get; set; }
        public string search { get; set; }
        public string key { get; set; }
        public int column { get; set; }
        public int row { get; set; }
        public string actionType { get; set; }
        public string status {  get; set; }
        public string actual { get; set; }
        public string expected { get; set; }
    }
}
