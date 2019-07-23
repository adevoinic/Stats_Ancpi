using System.IO;
using System.Collections.Generic;

namespace Data
{
	public class CsvTool
	{
		public void ExportToCsv(IEnumerable<TranzactiiXls> tranzactiis, string fileName) {
			if (File.Exists(fileName)) {
				File.Delete(fileName);
			}

			using (var fs = File.CreateText(fileName)) {

				fs.WriteLine(TranzactiiXls.CsvHeader);
				foreach (var trans in tranzactiis)
				{
					 
					fs.WriteLine(trans.CsvValues);
				}				
			}
		}
	}
}
