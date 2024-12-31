using Google.Apis.Sheets.v4.Data;
using GoogleSheetsHelper;
using Microsoft.Extensions.DependencyInjection;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services.Tests
{
	public class RequestCreatorTests : SheetGeneratorTests
	{
		// NOTE: Unless otherwise specified, all tests here assume an 8-team division

		private const int SHEET_ID = 12345;
		private const int GAMES_PER_ROUND = 2;
		private const int START_ROW_IDX = 2;
		private const int TEAMS_PER_POOL = 4;

		[Fact]
		public void CanCreateHeaderRequests()
		{
			DivisionSheetConfig config = DivisionSheetConfigFactory.GetForTeams(8);
			List<Team> teams = CreateTeams(config);
			PsoDivisionSheetHelper helper = new PsoDivisionSheetHelper(config);
			ScoreSheetHeadersRequestCreator creator = new ScoreSheetHeadersRequestCreator(config, helper);

			PoolPlayInfo info = new PoolPlayInfo(teams);
			IEnumerable<Team> poolTeams = info.Pools!.First();
			info = creator.CreateHeaderRequests(info, POOL_A, START_ROW_IDX, poolTeams);

			Assert.All(info.UpdateValuesRequests, rq => Assert.Single(rq.Rows));
			Assert.Collection(info.UpdateValuesRequests
				, rq =>
				{
					// pool label
					Assert.All(rq.Rows.Single(), cell => Assert.True(cell.GoogleBackgroundColor.GoogleColorEquals(Colors.HeaderRowColor)));
					GoogleSheetCell? labelCell = rq.Rows.Single().SingleOrDefault(cell => !string.IsNullOrEmpty(cell.StringValue));
					Assert.NotNull(labelCell);
					Assert.Equal($"POOL {POOL_A}", labelCell!.StringValue);
					Assert.Equal(START_ROW_IDX, rq.RowStart);
				}
				, rq =>
				{
					// standings headers
					IEnumerable<string> headerValues = rq.Rows.Single().Select(x => x.StringValue);
					headerValues.Should().BeEquivalentTo(PsoDivisionSheetHelper.StandingsHeaderRow);
					Assert.All(rq.Rows.Single(), cell => Assert.True(cell.Bold));
					Assert.Equal(START_ROW_IDX + 1, rq.RowStart);
					Assert.Equal(helper.GetColumnIndexByHeader(Constants.HDR_RANK), rq.ColumnStart);
				}
				, rq =>
				{
					// tiebreakers headers
					Assert.Equal(poolTeams.Count() + PsoDivisionSheetHelper.MainTiebreakerColumns.Count, rq.Rows.Single().Count());
					IEnumerable<string> headerValues = rq.Rows.Single().Where(x => x.FormulaValue != null).Select(x => x.FormulaValue);
					Assert.All(headerValues, hdr => hdr.StartsWith($"=\"vs \" & LEFT({ShootoutConstants.SHOOTOUT_SHEET_NAME}!{POOL_A}"));
					Assert.Equal(START_ROW_IDX + 1, rq.RowStart);
					Assert.Equal(helper.GetColumnIndexByHeader(Constants.HDR_GOAL_DIFF) + 1, rq.ColumnStart);
				}
			);

			Assert.Single(info.UpdateSheetRequests); // for the tiebreaker note
			Request noteRequest = info.UpdateSheetRequests.Single();
			Assert.NotNull(noteRequest.UpdateCells);
			Assert.Equal(START_ROW_IDX + 1, noteRequest.UpdateCells.Range.StartRowIndex);
			Assert.Equal(helper.GetColumnIndexByHeader(Constants.HDR_TIEBREAKER_H2H), noteRequest.UpdateCells.Range.StartColumnIndex);
		}

		private void ValidateGameRanges(Request request, DivisionSheetConfig config)
		{
			for (int i = 0; i < config.GamesPerRound; i++)
			{
				int startRow = START_ROW_IDX + 1 + (i * config.GamesPerRound) + (i == 0 ? 0 : i * Constants.ROUND_OFFSET_STANDINGS_TABLE);
				int endRow = startRow + config.GamesPerRound - 1;
				string cellRange = Utilities.CreateCellRangeString("A", startRow, endRow, CellRangeOptions.FixRow);
				Assert.Contains(cellRange, request.RepeatCell.Cell.UserEnteredValue.FormulaValue);
			}
		}

		[Theory]
		[InlineData(10)] // good for 5-and 10-team divisions
		[InlineData(8)] // good for all other divisions
		public void CanCreateStandingsRequests(int numTeams)
		{
			DivisionSheetConfig config = DivisionSheetConfigFactory.GetForTeams(numTeams);
			List<Team> teams = CreateTeams(config);
			PsoDivisionSheetHelper helper = numTeams == 10 ? new PsoDivisionSheetHelper10Teams(config) : new PsoDivisionSheetHelper(config);
			StandingsTableRequestCreator creator = new StandingsTableRequestCreator(config, helper, CreateStandingsRequestCreatorFactory(teams, config));

			PoolPlayInfo info = new PoolPlayInfo(teams);
			info = creator.CreateStandingsRequests(info, teams, START_ROW_IDX);
			IEnumerable<Request> standingsFormulaRequests = info.UpdateSheetRequests.Where(x => x.RepeatCell != null && x.RepeatCell.Range.StartColumnIndex <= helper.GetColumnIndexByHeader(Constants.HDR_GOAL_DIFF));

			// confirm that we have a request for each column that has a formula
			foreach (string hdr in helper.StandingsTableColumns)
			{
				int idx = helper.GetColumnIndexByHeader(hdr);
				Request? rq = standingsFormulaRequests.SingleOrDefault(x => x.RepeatCell.Range.StartColumnIndex == idx);
				switch (hdr)
				{
					case Constants.HDR_YELLOW_CARDS:
					case Constants.HDR_RED_CARDS:
					case ShootoutConstants.HDR_OVERALL_RANK:
						Assert.Null(rq);
						continue;
					default:
						Assert.NotNull(rq);
						break;
				}
			}

			// for games played/wins/draws/losses columns, confirm the score input ranges
			IEnumerable<Request> scoreBasedRequests = info.UpdateSheetRequests.Where(x => x.RepeatCell != null && x.RepeatCell.Range.StartColumnIndex > helper.GetColumnIndexByHeader(Constants.HDR_TEAM_NAME) && x.RepeatCell.Range.StartColumnIndex < helper.GetColumnIndexByHeader(Constants.HDR_YELLOW_CARDS));
			Assert.All(scoreBasedRequests, rq => ValidateGameRanges(rq, config));

			// confirm the columns and game ranges for the head-to-head tiebreakers
			IEnumerable<Request> headToHeadRequests = info.UpdateSheetRequests.Where(x => x.RepeatCell != null && x.RepeatCell.Range.StartColumnIndex > helper.GetColumnIndexByHeader(Constants.HDR_GOAL_DIFF));
			Assert.Equal(config.TeamsPerPool, headToHeadRequests.Count());
			int lastGameRowNum = (numTeams % 5) == 0 ? 20 : 12;
			string expectedCellRange = Utilities.CreateCellRangeString("A", 3, lastGameRowNum, CellRangeOptions.FixRow);
			Assert.All(headToHeadRequests, rq => Assert.Contains(expectedCellRange, rq.RepeatCell.Cell.UserEnteredValue.FormulaValue));

			// confirm the column indexes for the tiebreaker columns
			int? maxStandingsColIdx = standingsFormulaRequests.Max(x => x.RepeatCell.Range.StartColumnIndex);
			Assert.NotNull(maxStandingsColIdx);
			int? minHeadToHeadColIdx = headToHeadRequests.Min(x => x.RepeatCell.Range.StartColumnIndex);
			Assert.NotNull(minHeadToHeadColIdx);
			Assert.Equal(maxStandingsColIdx!.Value + 1, minHeadToHeadColIdx!.Value);
			int? maxHeadToHeadColIdx = headToHeadRequests.Max(x => x.RepeatCell.Range.StartColumnIndex);
			Assert.NotNull(maxHeadToHeadColIdx);
			Assert.Equal(maxStandingsColIdx!.Value + config.TeamsPerPool, maxHeadToHeadColIdx!.Value);
		}

		[Fact]
		public void CanCreateScoringRequests()
		{
			// this test validates the requests for creating the scoring section of the sheet; this method is executed once per round of games.
			// there should be 4 update sheet requests (formulas for team names and points) and 1 update values request (winners/pts headers, which get done here and not with the rest of the headers)
			DivisionSheetConfig config = DivisionSheetConfigFactory.GetForTeams(8);
			List<Team> teams = CreateTeams(config);
			PsoDivisionSheetHelper helper = new PsoDivisionSheetHelper(config);
			ScoreInputsRequestCreator creator = new ScoreInputsRequestCreator(config, helper, CreateStandingsRequestCreatorFactory(teams, config));

			PoolPlayInfo info = new PoolPlayInfo(teams);
			int startRowIndex = START_ROW_IDX;
			info = creator.CreateScoringRequests(info, teams, 1, ref startRowIndex);

			Assert.Equal(2, info.UpdateValuesRequests.Count);
			GoogleSheetRow roundLabelRow = info.UpdateValuesRequests.First().Rows.Single();
			Assert.Equal(PsoDivisionSheetHelper.GameScoreColumns.Count, roundLabelRow.Count);
			Assert.Equal("ROUND 1", roundLabelRow.First().StringValue);
			Assert.Equal(Constants.HDR_FORFEIT, roundLabelRow.Last().StringValue);
			Assert.All(roundLabelRow, cell => Assert.True(cell.GoogleBackgroundColor.GoogleColorEquals(Colors.SubheaderRowColor)));

			Assert.Single(info.UpdateValuesRequests.Last().Rows);
			UpdateRequest winnersAndPtsHeaderRequest = info.UpdateValuesRequests.Last();
			IEnumerable<string> headers = winnersAndPtsHeaderRequest.Rows.Single().Select(x => x.StringValue);
			headers.Should().BeEquivalentTo(PsoDivisionSheetHelper.WinnerAndPointsColumns);
			Assert.Equal(helper.GetColumnIndexByHeader(PsoDivisionSheetHelper.WinnerAndPointsColumns.First()), winnersAndPtsHeaderRequest.ColumnStart);

			// the first two are data validation requests for the home/away team inputs
			int startGamesRowIdx = START_ROW_IDX + 1;
			IEnumerable<Request> dvRequests = info.UpdateSheetRequests.Where(r => r.SetDataValidation != null);
			Assert.Equal(2, dvRequests.Count());
			Assert.All(dvRequests, r => Assert.Equal(startGamesRowIdx, r.SetDataValidation.Range.StartRowIndex));
			Assert.All(dvRequests, r => Assert.Equal(startGamesRowIdx + config.GamesPerRound, r.SetDataValidation.Range.EndRowIndex));

			// then we have the forfeit checkbox request and the winner and points column requests
			IEnumerable<Request> repeatRequests = info.UpdateSheetRequests.Where(r => r.RepeatCell != null);
			Assert.Equal(PsoDivisionSheetHelper.WinnerAndPointsColumns.Count + 1, repeatRequests.Count());
			Assert.All(repeatRequests, r => Assert.Equal(startGamesRowIdx, r.RepeatCell.Range.StartRowIndex));
			Assert.All(repeatRequests, r => Assert.Equal(startGamesRowIdx + config.GamesPerRound, r.RepeatCell.Range.EndRowIndex));

			// the rest are the resize requests
			IEnumerable<Request> resizeRequests = info.UpdateSheetRequests.Where(r => r.UpdateDimensionProperties != null);
			Assert.Equal(helper.StandingsTableColumns.Count(), resizeRequests.Count());
			Assert.All(resizeRequests, r => Assert.NotEqual(-1, r.UpdateDimensionProperties.Range.StartIndex));

			// startRowIndex should be ready for the next round
			Assert.Equal(START_ROW_IDX + GAMES_PER_ROUND + 2, startRowIndex); // +1 for the round label, +1 for a blank space after
		}

		/// <summary>
		/// This test validates the creation of the requests for creating conditional formatting to highlight the winners
		/// as well as the request to create the legend.
		/// </summary>
		[Fact]
		public void CanCreateWinnerFormattingRequests()
		{
			// expect 1 values request to create the legend and 5 update requests (4 conditional formatting to highlight the winners and 1 to right-align the legend text)

			DivisionSheetConfig config = DivisionSheetConfigFactory.GetForTeams(8);
			List<Team> teams = CreateTeams(config);
			PsoDivisionSheetHelper helper = new PsoDivisionSheetHelper(config);
			WinnerFormattingRequestsCreator creator = new WinnerFormattingRequestsCreator(config, new PsoFormulaGenerator(helper));

			Fixture fixture = new Fixture();
			fixture.Customize<PoolPlayInfo>(c => c.FromFactory(() => new PoolPlayInfo(teams))); // https://stackoverflow.com/questions/26149618/autofixture-customizations-provide-constructor-parameter
			PoolPlayInfo poolInfo = fixture.Build<PoolPlayInfo>()
				.Without(x => x.UpdateSheetRequests)
				.Without(x => x.UpdateValuesRequests)
				.With(x => x.StandingsStartAndEndRowNums, fixture.CreateMany<Tuple<int, int>>(2).ToList())
				.Create();
			ChampionshipInfo info = new ChampionshipInfo(poolInfo);

			SheetRequests requests = creator.CreateWinnerFormattingRequests(info);

			// verify the legend
			Assert.Single(requests.UpdateValuesRequests);
			Assert.Equal(4, requests.UpdateValuesRequests.Single().Rows.Count); // one row per rank
			Assert.All(requests.UpdateValuesRequests.Single().Rows, row => Assert.Equal(2, row.Count)); // 2 cells per row: one for the ordinal value and one for the color
			Action<GoogleSheetCell, int> assertLegendColor = (cell, rank) =>
			{
				Assert.NotNull(cell.GoogleBackgroundColor);
				Assert.True(cell.GoogleBackgroundColor.GoogleColorEquals(Colors.GetColorForRank(rank)));
			};
			Assert.Collection(requests.UpdateValuesRequests.Single().Rows
				, row => assertLegendColor(row.Last(), 1)
				, row => assertLegendColor(row.Last(), 2)
				, row => assertLegendColor(row.Last(), 3)
				, row => assertLegendColor(row.Last(), 4)
				);

			Assert.Equal(5, requests.UpdateSheetRequests.Count);

			// verify the conditional format requests
			Action<GridRange, Tuple<int, int>> assertRange = (rg, pair) =>
			{
				Assert.Equal(pair.Item1 - 1, rg.StartRowIndex);
				Assert.Equal(pair.Item2, rg.EndRowIndex);
				Assert.Equal(helper.GetColumnIndexByHeader(helper.StandingsTableColumns.First()), rg.StartColumnIndex);
				Assert.Equal(helper.GetColumnIndexByHeader(helper.StandingsTableColumns.Last()) + 1, rg.EndColumnIndex);
			};
			Action<AddConditionalFormatRuleRequest, int> assertFormatRequest = (rq, rank) =>
			{
				int gameRowNum = rank <= 2 ? info.ChampionshipGameRowNum : info.ThirdPlaceGameRowNum;
				string formula = rq.Rule.BooleanRule.Condition.Values.Single().UserEnteredValue;
				Assert.Contains($"${helper.HomeTeamColumnName}${gameRowNum}", formula);
				Assert.Contains($"${helper.TeamNameColumnName}{info.StandingsStartAndEndRowNums.First().Item1}", formula);
				Assert.True(Colors.GetColorForRank(rank).GoogleColorEquals(rq.Rule.BooleanRule.Format.BackgroundColor));
				Assert.Collection(rq.Rule.Ranges
					, rg => assertRange(rg, info.StandingsStartAndEndRowNums.First())
					, rg => assertRange(rg, info.StandingsStartAndEndRowNums.Last())
				);
			};
			IEnumerable<Request> formatRequests = requests.UpdateSheetRequests.Where(x => x.AddConditionalFormatRule != null);
			Assert.Collection(formatRequests
				, rq => assertFormatRequest(rq.AddConditionalFormatRule, 1)
				, rq => assertFormatRequest(rq.AddConditionalFormatRule, 2)
				, rq => assertFormatRequest(rq.AddConditionalFormatRule, 3)
				, rq => assertFormatRequest(rq.AddConditionalFormatRule, 4)
				);

			// verify the resize request
			Assert.NotNull(requests.UpdateSheetRequests.Last().RepeatCell);
		}

		[Fact]
		public void PoolPlayCreatorCanSetChampionshipStartRowIndex()
		{
			DivisionSheetConfig config = DivisionSheetConfigFactory.GetForTeams(8);
			List<Team> teams = CreateTeams(config);
			PoolPlayInfo info = new PoolPlayInfo(teams);
			CreateMocksForPoolPlayRequestCreatorTests(config);

			PoolPlayRequestCreator creator = new(config, _mockHeadersCreator.Object, _mockInputsCreator.Object, _mockStandingsCreator.Object, _mockTiebreakerColsCreator.Object, _mockSortedStandingsCreator.Object);
			info = creator.CreatePoolPlayRequests(info);

			Assert.Equal(26, info.ChampionshipStartRowIndex); // this number comes from the 2021 score sheet for 10U Boys, as that was an 8-team division
		}

		[Theory]
		[InlineData(8)]
		[InlineData(6)]
		[InlineData(5)]
		[InlineData(4)]
		public void PoolPlayCreatorCanSetStandingsStartAndEndRows(int numTeams)
		{
			DivisionSheetConfig config = DivisionSheetConfigFactory.GetForTeams(numTeams);
			List<Team> teams = CreateTeams(config);
			PoolPlayInfo info = new PoolPlayInfo(teams);
			CreateMocksForPoolPlayRequestCreatorTests(config);

			PoolPlayRequestCreator creator = new(config, _mockHeadersCreator.Object, _mockInputsCreator.Object, _mockStandingsCreator.Object, _mockTiebreakerColsCreator.Object, _mockSortedStandingsCreator.Object);
			info = creator.CreatePoolPlayRequests(info);

			Assert.Equal(numTeams <= 5 ? 1 : 2, info.StandingsStartAndEndRowNums.Count);
		}
	}
}
