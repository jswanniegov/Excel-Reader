using SpreadsheetReader.models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace SpreadsheetReader.models
{
    /* 
    * Custom Class
    * The class is a model of the vrcTest table used to store excel data
    */
    public class RowItem
    {
        public List<Column> columnList = new List<Column>(); 
        public string vrcRefNum { set; get; }
        public string vrcTaxType { set; get; }
        public string vrcPeriod { set; get; }
        public string vrcTaxYear { set; get; }
        public string vrcRiskDescription { set; get; }
        public string vrcRiskID { set; get; }
        public string vrcCaseType { set; get; }
        public string vrcSuppressCommunicationInd { set; get; }
        public string vrcCaseNum { set; get; }
        public string vrcRequestOperation { set; get; }
        public string vrcComments { set; get; }

        public string SQL()
        {
            return "('" + this.vrcRefNum + "', '"
                    + this.vrcTaxType + "', '"
                    + this.vrcPeriod + "', '"
                    + this.vrcTaxYear + "', '"
                    + this.vrcRiskDescription + "', '"
                    + this.vrcRiskDescription + "', '"
                    + this.vrcRiskID + "', '"
                    + this.vrcSuppressCommunicationInd + "', '"
                    + this.vrcCaseNum + "', '"
                    + this.vrcRequestOperation + "', '"
                    + this.vrcComments + "'" + ")";
        }


        public string ReusableSQL()
        {
            string sql = "("; 
            for (int i = 0; i < columnList.Count; i++)
            {
                if(i == (columnList.Count - 1))
                    sql = "'" + columnList[i] + "')";
                else
                    sql = "'" + columnList[i] + "', ";
            }
            return sql; 
        }
    }
}