namespace PantherShootoutScoreSheetGenerator.Services
{
	public class ShootoutSheetConfig
	{
		public int? SheetId { get; set; }
		/// <summary>
		/// Maximum cell width for team names, based on the team with the longest name
		/// </summary>
		public int TeamNameCellWidth { get; set; }
		/// <summary>
		/// Holds the <see cref="DivisionSheetConfig"/> for each division
		/// </summary>
		public Dictionary<string, DivisionSheetConfig>? DivisionConfigs { get; set; }
		/// <summary>
		/// Holds the start and end row numbers for shootout score entries
		/// </summary>
		public Dictionary<string, Tuple<int, int>> ShootoutStartAndEndRows { get; set; } = new();
	}
}
