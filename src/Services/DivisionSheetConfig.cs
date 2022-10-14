namespace PantherShootoutScoreSheetGenerator.Services
{
	public class DivisionSheetConfig
	{
		public int? SheetId { get; set; }
		public int GamesPerRound { get; set; }
		public int NumberOfRounds { get; set; }
		public int TeamNameCellWidth { get; set; }
		public int TeamsPerPool { get; set; }
	}
}
