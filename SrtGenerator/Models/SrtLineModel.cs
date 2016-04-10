using System;
using System.Collections.Generic;
using System.Text;

namespace SrtGenerator.Models
{
    public class SrtLineModel
    {
        public int Line { get; set; }
        public DateTime? StartDate { get; set; }
        public DateTime? EndDate { get; set; }
        public Dictionary<String, String> Translations { get; set; }   
    }
}
