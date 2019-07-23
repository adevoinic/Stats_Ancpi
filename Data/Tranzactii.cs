using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Data
{
	public class TranzactiiXls
	{
		public int IdJudet { get; set; }

		public string NumeJudet { get; set; }

		public DateTime Data { get; set; }

		public int Teren_Extra_Agricol { get; set; }
		
		public int Teren_Extra_NeAgricol { get; set; }
		
		public int Teren_Intra_CuConstructii { get; set; }
		
		public int Teren_Intra_FaraConstructii { get; set; }

		public int UnitatiIndividuale { get; set; }		

		public static string CsvHeader {
			get
			{
				return string.Join(",", "IdJudet", "NumeJudet", "Data", "Teren_Extra_Agricol", "Teren_Extra_NeAgricol",
										"Teren_Intra_CuConstructii", "Teren_Intra_FaraConstructii", "UnitatiIndividuale");
			}
		}

		public string CsvValues
		{
			get
			{
				return string.Join(",", IdJudet, NumeJudet, Data.ToShortDateString(), Teren_Extra_Agricol, Teren_Extra_NeAgricol,
							   Teren_Intra_CuConstructii, Teren_Intra_FaraConstructii, UnitatiIndividuale);
			}
		}
	}
}
