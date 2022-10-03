using LINQtoCSV;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public class Team
	{
		[CsvColumn(FieldIndex = 1)]
		public string DivisionName { get; set; } = string.Empty;
		[CsvColumn(FieldIndex = 2)]
		public string TeamName { get; set; } = string.Empty;
		public string PoolName { get; set; } = string.Empty;
		public string TeamSheetCell { get; set; } = string.Empty;
	}
}
