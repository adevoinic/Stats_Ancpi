using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Data
{
	public class ExcelParser
	{
		private IList<string> valuableFiles = new List<string>();

		public ExcelParser()
		{
			valuableFiles.Add("iunie_2019_vanzari.xlsx");
			valuableFiles.Add("Mai_2019_Vanzari.xlsx");
			valuableFiles.Add("Aprilie_vanzari_2019.xlsx");
			valuableFiles.Add("Martie_vanzari_2019.xlsx");
			valuableFiles.Add("20190301_vanzari.xlsx"); //februarie 2019
			valuableFiles.Add("Ianuarie_vanzari_2019.xlsx");
			//-----------------       2018
			valuableFiles.Add("Decembrie_vanzari_2018.xlsx");
			valuableFiles.Add("Noiembrie_vanzari_2018.xlsx");
			valuableFiles.Add("Octombrie_vanzari_2018.xlsx");
			valuableFiles.Add("Septembrie_vanzari_2018.xlsx");

			valuableFiles.Add("dinamica_August_tranzactii_2018.xlsx");
			valuableFiles.Add("20180801_vanzari_site_e3.xlsx");			 
			valuableFiles.Add("6_iunie_vanzari_site_e3.xlsx");
			valuableFiles.Add("dinamica_vanzari_mai_2018.xlsx");
			
			valuableFiles.Add("20180501_vanzari_site_e3.xlsx");
			valuableFiles.Add("20180401_vanzari_site_e3.xlsx");
			valuableFiles.Add("20180301_vanzari_site_e3.xlsx");
			valuableFiles.Add("DINAMICA_VANZARI_IANUARIE_2018.xlsx");

			/// 2017		 
			valuableFiles.Add("20170701_vanzari_site_e3.xlsx");
			valuableFiles.Add("20170801_vanzari_site_e3.xlsx");
			valuableFiles.Add("20170901_vanzari_site_e3.xlsx");
			valuableFiles.Add("20171001_vanzari_site_e3_2.xlsx");

			valuableFiles.Add("Vanzari_octombrie_2017.xlsx");
			valuableFiles.Add("Vanzari_noiembrie_2017.xlsx");
			valuableFiles.Add("VANZARI_DECEMBRIE_2017.xlsx");			
		}

		public List<TranzactiiXls> ImportExcelData()
		{
			List<TranzactiiXls> transactions = new List<TranzactiiXls>();
			var files = Directory.GetFiles(DataDownloader.DownloadFolder);
			foreach (var file in files)
			{
				if (!file.EndsWith(DataDownloader.XlsxExtension))
				{
					continue;
				}
				if (valuableFiles.FirstOrDefault(f => file.Contains(f)) == null)
				{
					continue;
				}
				transactions.AddRange(ReadXlsFile(file));
			}
			return transactions;
		}

		private IList<TranzactiiXls> ReadXlsFile(string file)
		{
			var fullPath = Environment.CurrentDirectory + "\\" + file;
			var excel = new Application();
			IList<TranzactiiXls> transactions = null;
			
			try
			{
				var workbook = excel.Workbooks.Open(fullPath);
				var sheet = workbook.Sheets.Item[1];
				Range excelRange = sheet.UsedRange;

				object[,] valueArray = (object[,])excelRange.get_Value(XlRangeValueDataType.xlRangeValueDefault);

				transactions = ExtractValues(valueArray);
				workbook.Close(SaveChanges:false);
			}
			catch (Exception ex)
			{
				throw;
			}
			finally
			{
				excel.Workbooks.Close();
			}
			return transactions;
		}
		
		private IList<TranzactiiXls> ExtractValues(object[,] valueArray)
		{
			var date = GetDateFromFile(valueArray);

			IList<TranzactiiXls> transactions = new List<TranzactiiXls>();
			var linesCount = valueArray.GetUpperBound(0);
			for (int line = 6; line < linesCount; line++)
			{
				if (valueArray[line, 1] == null)
				{
					continue;
				}
				var transaction = new TranzactiiXls();
				transaction.IdJudet = int.Parse(valueArray[line, 1].ToString());
				transaction.NumeJudet = valueArray[line, 2].ToString();
				transaction.Teren_Extra_Agricol = int.Parse(valueArray[line, 3].ToString());
				transaction.Teren_Extra_NeAgricol = int.Parse(valueArray[line, 4].ToString());
				transaction.Teren_Intra_CuConstructii = int.Parse(valueArray[line, 5].ToString());
				transaction.Teren_Intra_FaraConstructii = int.Parse(valueArray[line, 6].ToString());
				transaction.UnitatiIndividuale = int.Parse(valueArray[line, 7].ToString());
				transaction.Data = date;
				transactions.Add(transaction);
			}
			return transactions;
		}

		private DateTime GetDateFromFile(object[,] valueArray)
		{
			var valueDate = valueArray[1, 1].ToString();
			int year = int.Parse(valueDate.Substring(valueDate.Length - 4));

			var index = valueDate.IndexOf("-");
			var spaceIndex = valueDate.IndexOf(" ", index + 2);

			var valueMonth = valueDate.Substring(index + 2, spaceIndex - index - 2);
			var month = GetMonth(valueMonth);
			return new DateTime(year, month, 28);
		}

		private int GetMonth(string valueMonth)
		{
			switch (valueMonth.ToLower())
			{
				case "ianuarie": return 1;
				case "februarie": return 2;
				case "martie": return 3;
				case "aprilie": return 4;
				case "mai": return 5;
				case "iunie": return 6;
				case "iulie": return 7;
				case "august": return 8;
				case "septembrie": return 9;
				case "octombrie": return 10;
				case "noiembrie": return 11;
				case "decembrie": return 12;
				default:
					throw new ApplicationException("Check date parser.");
			}
		}
	}
}
