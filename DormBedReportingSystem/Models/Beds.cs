using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace DormBedReportingSystem.Models
{
    public class Beds
    {
        public int id { get; set; }

        public string REGION { get; set; }

        public string REGION_NAME { get; set; }

        public string CENTER_ID { get; set; }

        public string CENTER_NAME { get; set; }

        public string BLDG_NAME { get; set; }

        public string WING_NAME { get; set; }
        [Required]
        public int CAP_1BED { get; set; }
        [Required]
        public int CAP_2BEDS { get; set; }
        [Required]
        public int CAP_3BEDS { get; set; }
        [Required]
        public int CAP_4BEDS { get; set; }
        [Required]
        public int CAP_5BEDS { get; set; }
        [Required]
        public int CAP_6BEDS { get; set; }
        [Required]
        public int CAP_7BEDS { get; set; }
        [Required]
        public int CAP_8BEDS { get; set; }
        [Required]
        public int CAP_9BEDS { get; set; }
        [Required]
        public int CAP_10BEDS { get; set; }

        public int CAP_TOTAL { get; set; }
        [Required]
        public int OCC_1BED { get; set; }
        [Required]
        public int OCC_2BEDS { get; set; }
        [Required]
        public int OCC_3BEDS { get; set; }
        [Required]
        public int OCC_4BEDS { get; set; }
        [Required]
        public int OCC_5BEDS { get; set; }
        [Required]
        public int OCC_6BEDS { get; set; }
        [Required]
        public int OCC_7BEDS { get; set; }
        [Required]
        public int OCC_8BEDS { get; set; }
        [Required]
        public int OCC_9BEDS { get; set; }
        [Required]
        public int OCC_10BEDS { get; set; }

        public int VB_OBS { get; set; }

        public int VB_REPAIR { get; set; }

        public int VB_NOFURNITURE { get; set; }

        public int OCC_TOTAL { get; set; }

        public string BedComments { get; set; }

    }
}
