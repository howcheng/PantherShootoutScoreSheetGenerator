namespace PantherShootoutScoreSheetGenerator.Services
{
	public class DivisionSheetConfig
	{
		public int? SheetId { get; set; }
		public int NumberOfTeams { get; set; }
		public int TeamsPerPool { get; set; }
		public int NumberOfPools => NumberOfTeams / TeamsPerPool;
		public int GamesPerRound { get; set; }
		public int NumberOfRounds { get; set; }
		public int TeamNameCellWidth { get; set; }

		public void SetupForTeams(int numTeams)
		{
			NumberOfTeams = numTeams;
			switch (numTeams)
			{
				case 6:
					NumberOfRounds = 3;
					GamesPerRound = 1;
					TeamsPerPool = 3;
					break;
				case 8:
					NumberOfRounds = 3;
					GamesPerRound = 2;
					TeamsPerPool = 4;
					break;
				case 10:
					NumberOfRounds = 5;
					GamesPerRound = 2;
					TeamsPerPool = 5;
					break;
				default: // 12 teams
					NumberOfRounds = 3;
					GamesPerRound = 2;
					TeamsPerPool = 4;
					break;
			}
		}
	}
}
