using GoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services.Tests
{
	public class ChampionshipRequestCreator8TeamsTests : SheetGeneratorTests
	{
		[Fact]
		public void CanCreateChampionshipRequests()
		{
			DivisionSheetConfig config = DivisionSheetConfigFactory.GetForTeams(8);
			PsoDivisionSheetHelper helper = new PsoDivisionSheetHelper(config);
			ChampionshipRequestCreator8Teams creator = new ChampionshipRequestCreator8Teams(ShootoutConstants.DIV_10UB, config, helper);

			List<Team> teams = CreateTeams(config);
			PoolPlayInfo info = new PoolPlayInfo(teams)
			{
				StandingsStartAndEndRowNums = new List<Tuple<int, int>>
				{
					// these row numbers from the 2021 score sheet for 10U Boys
					new Tuple<int, int>(3, 6),
					new Tuple<int, int>(16, 19)
				}
			};
			ChampionshipInfo result = creator.CreateChampionshipRequests(info);

			// expect 5 values updates: 1 for the label, 1 subheader and game input row each for the consolation and championship
			Action<string, int, Tuple<int, int>> assertChampionshipTeamFormula = (f, poolRank, items) =>
			{
				string expected = string.Format("=IF(COUNTIF(F{0}:F{1},\"=3\")=4, VLOOKUP({2},{{M{0}:M{1},E{0}:E{1}}},2,FALSE), \"\")", items.Item1, items.Item2, poolRank);
				Assert.Equal(expected, f);
			};
			Action<UpdateRequest, int, int> assertChampionshipScoreEntry = (rq, rank, rowIdx) =>
			{
				Assert.Equal(rowIdx, rq.RowStart);
				GoogleSheetRow row = rq.Rows.Single();
				Assert.Equal(4, row.Count()); // 4 cells
				Assert.Collection(row
					, cell => assertChampionshipTeamFormula(cell.FormulaValue, rank, info.StandingsStartAndEndRowNums.First())
					, cell => Assert.Empty(cell.StringValue)
					, cell => Assert.Empty(cell.StringValue)
					, cell => assertChampionshipTeamFormula(cell.FormulaValue, rank, info.StandingsStartAndEndRowNums.Last())
					);
			};
			int rowIndex = info.ChampionshipStartRowIndex;
			Assert.Collection(result.UpdateValuesRequests
				, rq => AssertChampionshipLabelRequest(rq, info, rowIndex)
				, rq => AssertChampionshipSubheader(rq, "3RD-PLACE", ++rowIndex)
				, rq => assertChampionshipScoreEntry(rq, 2, ++rowIndex)
				, rq => AssertChampionshipSubheader(rq, "CHAMPIONSHIP", ++rowIndex)
				, rq => assertChampionshipScoreEntry(rq, 1, ++rowIndex)
				);
		}
	}
}
