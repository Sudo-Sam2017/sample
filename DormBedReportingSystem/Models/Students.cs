using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace DormBedReportingSystem.Models
{


    public class Students
    {
        public string TOTAL_TARGET { get; set; } //contractStrength
        public string AS_OF_TARGET { get; set; }

        public string RES_STUD_FEMALE /*cs_femaleResidents*/ { get; set; }

        public int cs_femaleNonResidents { get; set; }

        public string RES_STUD_MALE { get; set; } //cs_maleResidents { get; set; }

        public string NRES_STUD_MALE { get; set; }

        public string TOTAL_MALE_TARGET { get; set; }

        public string TOTAL_FEM_TARGET { get; set; }

        public int cs_totalMaleResidents { get; set; }
        public int cs_totalFemaleResidents { get; set; }

        public int cs_totalMaleNonResidents { get; set; }

        public string NRES_STUD_FEMALE { get; set; }

        public string TOTAL_RES_TARGET /*cs_totalResidents*/ { get; set; }

        public string TOTAL_NRES_TARGET { get; set; }

        public string TOTAL_ACTUAL { get; set; } //on Board Strength

        public string AS_OF_ACTUAL { get; set; }

        public string RES_STUD_FEMALE_OBS { get; set; }

        public string NRES_STUD_FEMALE_OBS { get; set; }

        public string RES_STUD_MALE_OBS { get; set; }

        public string NRES_STUD_MALE_OBS { get; set; }

        public string TOTAL_FEM_ACTUAL { get; set; }

        public string TOTAL_MALE_ACTUAL { get; set; }

        public int obs_totalMaleResidents { get; set; }

        public int obs_totalFemaleResidents { get; set; }

        public int obs_totalMaleNonResidents { get; set; }

        public int obs_totalFemaleNonResidents { get; set; }

        public string TOTAL_RES_ACTUAL { get; set; } //obs

        public string TOTAL_NRES_ACTUAL { get; set; } //obs


    }

}
