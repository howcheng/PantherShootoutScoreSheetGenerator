using GoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services.Tests
{
	public class ChampionshipRequestCreator6TeamsTests : SheetGeneratorTests
	{
		[Fact]
		public void CanCreateChampionshipRequests()
		{
			DivisionSheetConfig config = DivisionSheetConfigFactory.GetForTeams(6);
			PsoDivisionSheetHelper helper = new PsoDivisionSheetHelper(config);
			ChampionshipRequestCreator6Teams creator = new ChampionshipRequestCreator6Teams(ShootoutConstants.DIV_10UB, config, helper);

			List<Team> teams = CreateTeams(config);
			PoolPlayInfo info = new PoolPlayInfo(teams) 
			{ 
				// these row numbers/indices taken from the 2021 score sheet for 14U Girls
				StandingsStartAndEndRowNums = new List<Tuple<int, int>>
				{
					new Tuple<int, int>(3, 5),
					new Tuple<int, int>(13, 15)
				},
				ChampionshipStartRowIndex = 20,
			};
			ChampionshipInfo result = creator.CreateChampionshipRequests(info);

			// expect 10 values updates: 1 for the label, 4 subheaders, 5 score entry rows
			Action<UpdateRequest, string, string, int> assertChampionshipScoreEntry = (rq, homeFormula, awayFormula, rowIdx) =>
			{
				Assert.Equal(rowIdx, rq.RowStart);
				GoogleSheetRow row = rq.Rows.Single();
				Assert.Equal(4, row.Count()); // 4 cells
				Assert.Collection(row
					, cell => Assert.Equal(homeFormula, cell.FormulaValue)
					, cell => Assert.Empty(cell.StringValue)
					, cell => Assert.Empty(cell.StringValue)
					, cell => Assert.Equal(awayFormula, cell.FormulaValue)
					);
			};

			int rowIndex = info.ChampionshipStartRowIndex;
			Assert.Collection(result.UpdateValuesRequests
				, rq => AssertChampionshipLabelRequest(rq, info, rowIndex)
				, rq => AssertChampionshipSubheader(rq, "CONSOLATION", ++rowIndex)
				, rq => assertChampionshipScoreEntry(rq, GetFirstRoundFormula(3, info.StandingsStartAndEndRowNums.First()), GetFirstRoundFormula(3, info.StandingsStartAndEndRowNums.Last()), ++rowIndex)
				, rq => AssertChampionshipSubheader(rq, "SEMIFINALS", ++rowIndex)
				, rq => assertChampionshipScoreEntry(rq, GetFirstRoundFormula(1, info.StandingsStartAndEndRowNums.First()), GetFirstRoundFormula(2, info.StandingsStartAndEndRowNums.Last()), ++rowIndex)
				, rq => assertChampionshipScoreEntry(rq, GetFirstRoundFormula(2, info.StandingsStartAndEndRowNums.First()), GetFirstRoundFormula(1, info.StandingsStartAndEndRowNums.Last()), ++rowIndex)
				, rq => AssertChampionshipSubheader(rq, "3RD-PLACE", ++rowIndex)
				, rq => assertChampionshipScoreEntry(rq, GetFinalFormula(26, false), GetFinalFormula(25, false), ++rowIndex) // as above, these row numbers come from the 2021 sheet for 14U Girls
				, rq => AssertChampionshipSubheader(rq, "FINAL", ++rowIndex)
				, rq => assertChampionshipScoreEntry(rq, GetFinalFormula(25, true), GetFinalFormula(26, true), ++rowIndex)
				);
		}

		private string GetFirstRoundFormula(int rank, Tuple<int, int> startAndEndRowNums)
			=> string.Format("=IF(COUNTIF(F{1}:F{2},\"=2\")=3, VLOOKUP({0},{{M{1}:M{2},E{1}:E{2}}},2,FALSE), \"\")", rank, startAndEndRowNums.Item1, startAndEndRowNums.Item2);

		private string GetFinalFormula(int rowNum, bool isFinal)
		{
			char op1 = isFinal ? '>' : '<';
			char op2 = isFinal ? '<' : '>';
			return string.Format("=IFNA(IFS(B{0}{1}C{0},A{0},B{0}{2}C{0},D{0}), \"\")", rowNum, op1, op2);
		}
	}
}
