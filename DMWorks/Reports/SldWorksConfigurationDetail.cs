using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMWorks
{
    public class SldWorksConfigurationDetail{

        private List<SldWorksFeatureDetail> featureDetails;
        
        public SldWorksConfigurationDetail(){
            featureDetails = new List<SldWorksFeatureDetail>();
            CenterOfMass_user = new double[3];
        }


        public string Name { get; set; }
        public string Material { get; set; }
        public double Mass_kg { get; set; }
        public double Volume_m3 { get; set; }
        public double[] CenterOfMass_m { get; set; }
        public double Mass_user { get; set; }
        public double Volume_user { get; set; }
        public double[] CenterOfMass_user { get; set; }
        public string Mass_user_note => Mass_user.ToString("N3");
        public string Volume_user_note => Volume_user.ToString("N3");
        public string CenterOfMass_user_note => CenterOfMass_user.Count() == 3 ? $"({CenterOfMass_user[0].ToString("N3")},{CenterOfMass_user[1].ToString("N3")},{CenterOfMass_user[2].ToString("N3")})" : "";

        public List<SldWorksFeatureDetail> FeatureDetails { get { return featureDetails; } set { featureDetails = value; } }



    }
}
