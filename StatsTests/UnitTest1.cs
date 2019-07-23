using Data;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;

namespace StatsTests
{
	[TestClass]
	public class UnitTest1
	{
		[TestMethod]
		public void GetData()
		{
			var dataDownloader = new DataDownloader();
			//dataDownloader.DownloadExcelData();			 
		}

		[TestMethod]
		public void ImportExcelData()
		{
			var excelParser = new ExcelParser();

			var transactions = excelParser.ImportExcelData();
			Assert.IsNotNull(transactions);
			Assert.IsTrue(transactions.Any());

			var bucurestiTransactions = transactions.Where(f => f.IdJudet == 10).OrderBy(f => f.Data);
			Assert.IsTrue(bucurestiTransactions.Any());

			var csvTool = new CsvTool();
			csvTool.ExportToCsv(bucurestiTransactions, @".\buc.csv");
		}
	}
}
