using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DormBedReportingSystem.Models;

namespace DormBedReportingSystem.ViewModels
{
    public class StudentBedViewModel
    {
        public List<Students> Students {get; set;}

        public List<Beds> Beds {get; set;}

        public List<ReadEditRow> ReadEditRow { get; set; }

        public AddedRow AddRow { get; set; }

    }
}
