namespace PantherShootoutScoreSheetGenerator.Services
{
	public class DivisionSheetConfig
	{
		public int? SheetId { get; set; }
		public string DivisionName { get; set; } = string.Empty;
		public int NumberOfTeams { get; set; }
		public int TeamsPerPool { get; set; }
		public int NumberOfPools => NumberOfTeams / TeamsPerPool;
		public int GamesPerRound { get; set; }
		public int TotalPoolPlayGames => TeamsPerPool - 1;
		public int NumberOfGameRounds { get; set; }
		public int NumberOfShootoutRounds { get; set; }
		/// <summary>
		/// Gets populated from <see cref="ShootoutSheetConfig.TeamNameCellWidth"/>
		/// </summary>
		public int TeamNameCellWidth { get; set; }

		protected void SetupForTeams(int numTeams)
		{
			NumberOfTeams = numTeams;
			switch (numTeams)
			{
				case 6:
					NumberOfGameRounds = NumberOfShootoutRounds = 3;
					GamesPerRound = 1;
					TeamsPerPool = 3;
					break;
				case 5:
				case 10:
					NumberOfGameRounds = 5;
					NumberOfShootoutRounds = 4;
					GamesPerRound = 2;
					TeamsPerPool = 5;
					break;
				default: // 4, 8, or 12 teams
					NumberOfGameRounds = NumberOfShootoutRounds = 3;
					GamesPerRound = 2;
					TeamsPerPool = 4;
					break;
			}
		}
	}

	public class DivisionSheetConfig4Teams : DivisionSheetConfig
	{
		public DivisionSheetConfig4Teams()
		{
			SetupForTeams(4);
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

	public class DivisionSheetConfig5Teams : DivisionSheetConfig
	{
		public DivisionSheetConfig5Teams()
		{
			SetupForTeams(5);
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
				case 4:
					return new DivisionSheetConfig4Teams();
				case 5:
					return new DivisionSheetConfig5Teams();
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
