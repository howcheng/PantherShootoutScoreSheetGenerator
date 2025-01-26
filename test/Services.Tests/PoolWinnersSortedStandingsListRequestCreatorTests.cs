using GoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services.Tests
{
	public class PoolWinnersSortedStandingsListRequestCreatorTests
	{
		private const int TEAMS_PER_POOL = 4;
		private const int GAMES_PER_ROUND = 2;
		private const int NUM_ROUNDS = 3;

		[Fact]
		public void CanCreateSortedStandingsListRequest()
		{
			// confirm that the cell ranges and sort column numbers are done correctly
			DivisionSheetConfig12Teams config = new() { GamesPerRound = GAMES_PER_ROUND, NumberOfGameRounds = NUM_ROUNDS, TeamsPerPool = TEAMS_PER_POOL };
			PsoDivisionSheetHelper12Teams helper = new(config);
			PsoFormulaGenerator fg = new(helper);

			PoolPlayInfo info = new()
			{
				StandingsStartAndEndRowNums = new()
				{
					new(3, 12),
					new(16, 25),
					new(29, 38),
				},
				ChampionshipStartRowIndex = 39, // normally this will be set by the PoolPlayRequestCreator
			};
			PoolWinnersSortedStandingsListRequestCreator creator = new(fg, config);
			info = creator.CreateSortedStandingsListRequest(info);

			Assert.Collection(info.UpdateValuesRequests,
				x =>
				{
					string formula = x.Rows.Last().Last().FormulaValue;
					Assert.Contains("{{AI3:BB3,AF42};{AI16:BB16,AF43};{AI29:BB29,AF44}}", formula);
					Assert.Contains("8,False", formula); // only need to check one; we can assume the rest are correct because of math
					Assert.DoesNotContain("16,False", formula);
				},
				x =>
				{
					string formula = x.Rows.Last().Last().FormulaValue;
					Assert.Contains("{{AI4:BB4,AF46};{AI17:BB17,AF47};{AI30:BB30,AF48}}", formula);
				});
		}
	}
}
