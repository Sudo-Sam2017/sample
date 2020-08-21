using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace DormBedReportingSystem.Models
{
    public class ReadEditRow
    {
        public int id { get; set; }
        public string REGION { get; set; }

        public string REGION_NAME { get; set; }

        public string CENTER_ID { get; set; }
        public string CENTER_NAME { get; set; }

        public string BLDG_NAME { get; set; }

        public string WING_NAME { get; set; }

        public string CAP_1BED { get; set; }

        public string CAP_2BEDS { get; set; }

        public string CAP_3BEDS { get; set; }

        public string CAP_4BEDS { get; set; }

        public string CAP_5BEDS { get; set; }

        public string CAP_6BEDS { get; set; }

        public string CAP_7BEDS { get; set; }

        public string CAP_8BEDS { get; set; }

        public string CAP_9BEDS { get; set; }

        public string CAP_10BEDS { get; set; }
        public string OCC_1BED { get; set; }

        public string OCC_2BEDS { get; set; }

        public string OCC_3BEDS { get; set; }

        public string OCC_4BEDS { get; set; }

        public string OCC_5BEDS { get; set; }

        public string OCC_6BEDS { get; set; }

        public string OCC_7BEDS { get; set; }

        public string OCC_8BEDS { get; set; }

        public string OCC_9BEDS { get; set; }

        public string OCC_10BEDS { get; set; }

        public string? BedComments { get; set; }

    }
}
