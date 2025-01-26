using GoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services.Tests
{
	public class SortedStandingsListRequestCreator10TeamsTests
	{
		private const int TEAMS_PER_POOL = 5;
		private const int START_ROW_NUM = 3;
		private const int GAMES_PER_ROUND = 2;
		private const int NUM_ROUNDS = 3;

		[Fact]
		public void CanCreateSortedStandingsListRequest()
		{
			// confirm that the cell ranges and sort column numbers are done correctly
			Fixture fixture = new();
			DivisionSheetConfig config = new DivisionSheetConfig { GamesPerRound = GAMES_PER_ROUND, NumberOfGameRounds = NUM_ROUNDS, TeamsPerPool = TEAMS_PER_POOL };
			PsoDivisionSheetHelper helper = new(config);
			PsoFormulaGenerator fg = new(helper);

			PoolPlayInfo info = new()
			{
				StandingsStartAndEndRowNums = new()
				{
					new(3, 7),
					new(24, 28),
				}
			};
			SortedStandingsListRequestCreator10Teams creator = new(fg, config);
			info = creator.CreateSortedStandingsListRequest(info);

			UpdateRequest request = info.UpdateValuesRequests.Last();
			string formula = request.Rows.Last().Last().FormulaValue;

			Assert.Contains("{G3:AB7;G24:AB28}", formula);
			Assert.Contains("8,False", formula); // only need to check one; we can assume the rest are correct because of math
		}
	}
}
