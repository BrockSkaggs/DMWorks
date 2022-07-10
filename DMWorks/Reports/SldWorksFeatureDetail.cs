using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMWorks
{
    public class SldWorksFeatureDetail{

        public SldWorksFeatureDetail(){

        }

        public string Name { get; set; }
        public DateTime CreationDate { get; set; }
        public DateTime ModifiedDate { get; set; }
        public string CreatedBy { get; set; }
        public string Description { get; set; }
        public bool IsSketch { get; set; }
        public string SketchStatus { get; set; }
    }
}
