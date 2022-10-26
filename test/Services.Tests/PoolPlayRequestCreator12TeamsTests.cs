using Google.Apis.Sheets.v4.Data;
using GoogleSheetsHelper;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services.Tests
{
	public class PoolPlayRequestCreator12TeamsTests : SheetGeneratorTests
	{
		[Fact]
		public void CanCreatePoolPlayRequestsFor12Teams()
		{
			DivisionSheetConfig config = DivisionSheetConfigFactory.GetForTeams(12);
			List<Team> teams = CreateTeams(config);
			PoolPlayInfo12Teams info = new PoolPlayInfo12Teams(teams);

			PsoDivisionSheetHelper12Teams helper = new PsoDivisionSheetHelper12Teams(config);
			PsoFormulaGenerator fg = new PsoFormulaGenerator(helper);
			var creators = CreateMocksForPoolPlayRequestCreatorTests(config);

			IStandingsRequestCreatorFactory factory = CreateStandingsRequestCreatorFactory(teams, config);
			PoolPlayRequestCreator12Teams creator = new PoolPlayRequestCreator12Teams(config, creators.Item1, creators.Item2, creators.Item3, fg, factory);
			info = (PoolPlayInfo12Teams)creator.CreatePoolPlayRequests(info);

			Action<UpdateRequest, List<string>, int> assertPoolWinnersHeaders = (rq, hdrs, rowIdx) =>
			{
				GoogleSheetRow row = rq.Rows.Single();
				Assert.Equal(rowIdx, rq.RowStart);
				Assert.Equal(helper.GetColumnIndexByHeader(PsoDivisionSheetHelper12Teams.PoolWinnersHeaderRow.First()), rq.ColumnStart);
				IEnumerable<string> headerValues = row.Select(x => x.StringValue);
				headerValues.Should().BeEquivalentTo(hdrs);
			};
			Action<string, Tuple<int, int>> assertMainFormulaRowNums = (formula, startAndEnd) =>
			{
				string cellRange = Utilities.CreateCellRangeString(helper.GetColumnNameByHeader(Constants.HDR_RANK), startAndEnd.Item1, startAndEnd.Item2);
				Assert.Contains(cellRange, formula);
			};
			Action<GoogleSheetRow, bool, int> assertMainFormulas = (row, isWinners, idx) =>
			{
				Assert.Equal(config.NumberOfPools, row.Count);
				Tuple<int, int> startAndEnd = info.StandingsStartAndEndRowNums.ElementAt(idx);
				Assert.All(row, cell => assertMainFormulaRowNums(cell.FormulaValue, startAndEnd)); // we don't need to validate the formulas themselves as that's done elsewhere
			};

			// expect to have 4 values requests (1 header row, 1 for formulas for team name/points/games played, for both winners and runners-up)
			int headerRowIdx = info.StandingsStartAndEndRowNums.First().Item1 - 2; // -1 for the header, then -1 to convert to index
			Assert.Collection(info.UpdateValuesRequests
				, rq => assertPoolWinnersHeaders(rq, PsoDivisionSheetHelper12Teams.PoolWinnersHeaderRow, headerRowIdx)
				, rq => assertPoolWinnersHeaders(rq, PsoDivisionSheetHelper12Teams.RunnersUpHeaderRow, headerRowIdx + config.NumberOfPools + 1)
				, rq => Assert.Collection(rq.Rows
					, row => assertMainFormulas(row, true, 0)
					, row => assertMainFormulas(row, true, 1)
					, row => assertMainFormulas(row, true, 2)
					)
				, rq => Assert.Collection(rq.Rows
					, row => assertMainFormulas(row, false, 0)
					, row => assertMainFormulas(row, false, 1)
					, row => assertMainFormulas(row, false, 2)
					)
				);

			// expect to have 6 sheet requests: 1 for the tiebreaker checkboxes and 2 each for the rank formulas (calculated and tiebreaker) for both winners and runners-up
			Assert.All(info.UpdateSheetRequests, rq => Assert.NotNull(rq.RepeatCell));

			// confirm that they all have the correct start row index value
			int winnersStartRowIdx = info.StandingsStartAndEndRowNums.First().Item1 - 1;
			int runnerUpStartRowIdx = winnersStartRowIdx + config.NumberOfPools + 1; // the runners-up section starts immediately after the winners section, +1 for the other header row

			IEnumerable<Request> checkboxRequests = info.UpdateSheetRequests.Where(rq => rq.RepeatCell.Cell.DataValidation != null);
			Action<GridRange, bool, int> assertRangeNumbers = (range, isWinners, colIdx) =>
			{
				int rowIdx = isWinners ? winnersStartRowIdx : runnerUpStartRowIdx;
				Assert.Equal(rowIdx, range.StartRowIndex);
				Assert.Equal(rowIdx + config.NumberOfPools, range.EndRowIndex);
				Assert.Equal(colIdx, range.StartColumnIndex);
			};
			Assert.Collection(checkboxRequests
				, rq => assertRangeNumbers(rq.RepeatCell.Range, true, helper.GetColumnIndexByHeader(ShootoutConstants.HDR_POOL_WINNER_TIEBREAKER))
				, rq => assertRangeNumbers(rq.RepeatCell.Range, false, helper.GetColumnIndexByHeader(ShootoutConstants.HDR_POOL_WINNER_TIEBREAKER))
				);

			IEnumerable<Request> rankRequests = info.UpdateSheetRequests.Except(checkboxRequests);
			Assert.Collection(rankRequests
				, rq => assertRangeNumbers(rq.RepeatCell.Range, true, helper.GetColumnIndexByHeader(ShootoutConstants.HDR_POOL_WINNER_RANK))
				, rq => assertRangeNumbers(rq.RepeatCell.Range, true, helper.GetColumnIndexByHeader(ShootoutConstants.HDR_POOL_WINNER_CALC_RANK))
				, rq => assertRangeNumbers(rq.RepeatCell.Range, false, helper.GetColumnIndexByHeader(ShootoutConstants.HDR_POOL_WINNER_RANK))
				, rq => assertRangeNumbers(rq.RepeatCell.Range, false, helper.GetColumnIndexByHeader(ShootoutConstants.HDR_POOL_WINNER_CALC_RANK))
				);
		}
	}
}
