using GoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services.Tests
{
	public class ChampionshipRequestCreator12TeamsTests : SheetGeneratorTests
	{
		[Fact]
		public void CanCreateChampionshipRequests()
		{
			DivisionSheetConfig config = DivisionSheetConfigFactory.GetForTeams(12);
			config.DivisionName = ShootoutConstants.DIV_10UB;
			PsoDivisionSheetHelper12Teams helper = new PsoDivisionSheetHelper12Teams(config);
			ChampionshipRequestCreator12Teams creator = new ChampionshipRequestCreator12Teams(config, helper);

			List<Team> teams = CreateTeams(config);
			PoolPlayInfo12Teams info = new PoolPlayInfo12Teams(teams) 
			{ 
				StandingsStartAndEndRowNums = new List<Tuple<int, int>>
				{
					new Tuple<int, int>(3, 6),
					new Tuple<int, int>(16, 19),
					new Tuple<int, int>(29, 32),
				},
				ChampionshipStartRowIndex = 39,
				PoolWinnersStartAndEndRowNums = new List<Tuple<int, int>>
				{ 
					new Tuple<int, int>(3, 5),
					new Tuple<int, int>(7, 9),
				}
			};
			ChampionshipInfo result = creator.CreateChampionshipRequests(info);

			// expect 5 values updates: 1 for the label, 1 subheader and game input row each for the consolation and championship
			Action<string, int, Tuple<int, int>> assertChampionshipTeamFormula = (f, poolRank, items) =>
			{
				string expected = string.Format("=IF(COUNTIF(AE{0}:AE{1}, 3)=3, VLOOKUP({2},{{AA{0}:AA{1},AC{0}:AC{1}}},2,FALSE), \"\")", items.Item1, items.Item2, poolRank);
				Assert.Equal(expected, f);
			};
			Action<UpdateRequest, int, int, int> assertChampionshipScoreEntry = (rq, homeRank, awayRank, rowIdx) =>
			{
				Assert.Equal(rowIdx, rq.RowStart);
				GoogleSheetRow row = rq.Rows.Single();
				Assert.Equal(4, row.Count()); // 4 cells
				Tuple<int, int> homeStartAndEnd = info.PoolWinnersStartAndEndRowNums.First();
				Tuple<int, int> awayStartAndEnd = homeRank < 3 ? info.PoolWinnersStartAndEndRowNums.First() : info.PoolWinnersStartAndEndRowNums.Last();
				Assert.Collection(row
					, cell => assertChampionshipTeamFormula(cell.FormulaValue, homeRank, homeStartAndEnd)
					, cell => Assert.Empty(cell.StringValue)
					, cell => Assert.Empty(cell.StringValue)
					, cell => assertChampionshipTeamFormula(cell.FormulaValue, awayRank, awayStartAndEnd)
					);
			};

			int rowIndex = info.ChampionshipStartRowIndex;
			Assert.Collection(result.UpdateValuesRequests
				, rq => AssertChampionshipLabelRequest(rq, info, rowIndex)
				, rq => AssertChampionshipSubheader(rq, "3RD-PLACE", ++rowIndex)
				, rq => assertChampionshipScoreEntry(rq, 3, 1, ++rowIndex)
				, rq => AssertChampionshipSubheader(rq, "CHAMPIONSHIP", ++rowIndex)
				, rq => assertChampionshipScoreEntry(rq, 1, 2, ++rowIndex)
				);
		}
	}
}
