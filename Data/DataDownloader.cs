using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;

namespace Data
{
	public class DataDownloader
	{
		//http://www.ancpi.ro/index.php/presa-3/statistici
		private string webAddress = "http://www.ancpi.ro/files/Presa/statistici/";
		internal const string DownloadFolder = "Excels";
		
		internal const string XlsxExtension = ".xlsx";
		 
		public void DownloadExcelData()
		{
			WebClient webClient = new WebClient();
			var content = webClient.DownloadString(webAddress);

			var decodedContent = WebUtility.HtmlDecode(content);
			var links = GetAllLinksFrom(decodedContent);

			var xlsLinks = links.Where(f => f.EndsWith(XlsxExtension));

			foreach (var link in xlsLinks)
			{
				webClient.DownloadFile(webAddress + link, DownloadFolder + @"\" + link);
			}
		}

		private IList<string> GetAllLinksFrom(string content)
		{
			var links = new List<string>();
			var start = 0;
			var href = @"<a href=""";
			var index = content.IndexOf(href, start, StringComparison.InvariantCultureIgnoreCase);
			var endHref = "\">";

			while (index > 0)
			{

				var lastIndex = content.IndexOf(endHref, index + href.Length);
				var link = content.Substring(index + href.Length, lastIndex - index - href.Length);

				links.Add(link);
				index = content.IndexOf(href, lastIndex);
			}

			return links;
		}
	}
}
