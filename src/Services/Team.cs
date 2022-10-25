using LINQtoCSV;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public class Team
	{
		private string? _poolName;

		[CsvColumn(FieldIndex = 1)]
		public string DivisionName { get; set; } = string.Empty;
		[CsvColumn(FieldIndex = 2)]
		public string TeamName { get; set; } = string.Empty;
		public string PoolName 
		{ 
			get => _poolName ?? TeamName.Substring(0, 1); 
			set => _poolName = value; 
		}
		/// <summary>
		/// Cell reference where the team's name is located; should NOT include the sheet name ("A3", not "Shootout!A3")
		/// </summary>
		public string TeamSheetCell { get; set; } = string.Empty;
	}
}
