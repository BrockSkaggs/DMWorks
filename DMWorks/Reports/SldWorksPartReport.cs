using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMWorks
{
    public class SldWorksPartReport: SldWorksFileReport{

        private List<SldWorksConfigurationDetail> configDetails;

        public SldWorksPartReport(){
            configDetails = new List<SldWorksConfigurationDetail>();
        }

        public List<SldWorksConfigurationDetail> ConfigDetails { get { return configDetails; } set { configDetails = value; } }



    }
}
