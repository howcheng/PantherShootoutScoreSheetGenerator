﻿namespace PantherShootoutScoreSheetGenerator.Services
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

		protected void SetupForTeams(int numTeams)
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

	public class DivisionSheetConfig6Teams : DivisionSheetConfig
	{
		public DivisionSheetConfig6Teams()
		{
			SetupForTeams(6);
		}
	}

	public class DivisionSheetConfig8Teams : DivisionSheetConfig
	{
		public DivisionSheetConfig8Teams()
		{
			SetupForTeams(8);
		}
	}

	public class DivisionSheetConfig10Teams : DivisionSheetConfig
	{
		public DivisionSheetConfig10Teams()
		{
			SetupForTeams(10);
		}
	}

	public class DivisionSheetConfig12Teams : DivisionSheetConfig
	{
		public DivisionSheetConfig12Teams()
		{
			SetupForTeams(12);
		}
	}

	public static class DivisionSheetConfigFactory
	{
		public static DivisionSheetConfig GetForTeams(int numTeams)
		{
			switch (numTeams)
			{
				case 6:
					return new DivisionSheetConfig6Teams();
				case 8:
					return new DivisionSheetConfig8Teams();
				case 10:
					return new DivisionSheetConfig10Teams();
				default:
					return new DivisionSheetConfig12Teams();
			}
		}
	}
}
