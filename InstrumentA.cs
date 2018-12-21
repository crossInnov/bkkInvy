using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;
using System.Threading;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;


public class InstrumentA :
{
        public string ValueDate;
        public string Derivative;
        public int Tenor;
        public string FinalMaturity;
        public string Strike;

	public InstrumentA(DataRow row)
    {
            this.ValueDate  = row["ValueDate"].ToString();
            this.Derivative = row["Derivative"].ToString();
            this.Tenor = Convert.ToInt32(row["Tenor"].ToString());
            this.FinalMaturity = row["FinalMaturity"].ToString();
            this.Strike = row["Strike"].ToString();
	}
}

       
 public void readExcelSheet()
 {
   string testPath = "C:\";
   string fileName = "babababa.xls";
   var dt = ReadExcelSheet(testPath + fileName, "Template$");
   List<DataRow> rows = (from row in dt.AsEnumerable() select row).ToList<DataRow>();
			
			
   InstrumentList = new List<InstrumentA>();
   foreach (var elem in rows)
            {
                var item = new InstrumentA(elem);
                InstrumentList.Add(item);
            }
}     
  
