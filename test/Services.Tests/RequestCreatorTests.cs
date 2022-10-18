using Google.Apis.Sheets.v4.Data;
using GoogleSheetsHelper;
using Microsoft.Extensions.DependencyInjection;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services.Tests
{
	public class RequestCreatorTests
	{
		private const string POOL_A = "A";
		private const string POOL_B = "B";
		private const int SHEET_ID = 12345;
		private const int GAMES_PER_ROUND = 2;
		private const int NUM_STANDINGS_FORMULAS = 11;
		private const int START_ROW_IDX = 2;
		private const int TEAMS_PER_POOL = 4;

		private DivisionSheetConfig CreateSheetConfig(int numRounds = 3) => new DivisionSheetConfig
		{
			SheetId = SHEET_ID,
			NumberOfRounds = numRounds,
			GamesPerRound = GAMES_PER_ROUND,
			TeamsPerPool = TEAMS_PER_POOL,
			TeamNameCellWidth = 100,
		};

		private static List<Team> CreateTeams(string pool = POOL_A)
		{
			int counter = 0;
			Fixture fixture = new Fixture();
			List<Team> teams = fixture.Build<Team>()
				.With(x => x.DivisionName, ShootoutConstants.DIV_10UB)
				.With(x => x.PoolName, pool)
				.With(x => x.TeamSheetCell, () => $"{pool}{++counter}")
				.CreateMany(TEAMS_PER_POOL)
				.ToList();
			return teams;
		}

		private static IStandingsRequestCreatorFactory CreateStandingsRequestCreatorFactory(IEnumerable<Team> teams, DivisionSheetConfig config)
		{
			IServiceCollection services = new ServiceCollection();
			services.AddPantherShootoutServices(teams, config);
			IServiceProvider provider = services.BuildServiceProvider();
			return provider.GetRequiredService<IStandingsRequestCreatorFactory>();
		}

		[Fact]
		public void CanCreateHeaderRequests()
		{
			List<Team> teams = CreateTeams();
			DivisionSheetConfig config = CreateSheetConfig();
			PsoDivisionSheetHelper helper = new PsoDivisionSheetHelper(config);
			ScoreSheetHeadersRequestCreator creator = new ScoreSheetHeadersRequestCreator(teams.First().DivisionName, helper);

			PoolPlayInfo info = new PoolPlayInfo(teams);
			info = creator.CreateHeaderRequests(info, POOL_A, START_ROW_IDX, teams);

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
					headerValues.Should().BeEquivalentTo(DivisionSheetCreator.StandingsHeaderRow);
					Assert.All(rq.Rows.Single(), cell => Assert.True(cell.Bold));
					Assert.Equal(START_ROW_IDX + 1, rq.RowStart);
					Assert.Equal(helper.GetColumnIndexByHeader(Constants.HDR_TEAM_NAME), rq.ColumnStart);
				}
				, rq =>
				{
					// tiebreakers headers
					Assert.Equal(teams.Count, rq.Rows.Single().Count());
					IEnumerable<string> headerValues = rq.Rows.Single().Select(x => x.FormulaValue);
					Assert.All(headerValues, hdr => hdr.StartsWith($"=\"vs \" & LEFT({ShootoutConstants.SHOOTOUT_SHEET_NAME}!{POOL_A}"));
					Assert.Equal(START_ROW_IDX + 1, rq.RowStart);
					Assert.Equal(helper.GetColumnIndexByHeader(Constants.HDR_GOAL_DIFF) + 1, rq.ColumnStart);
				}
			);

			Assert.Single(info.UpdateSheetRequests); // for the tiebreaker note
		}

		[Theory]
		[InlineData(5)] // 10-team division
		[InlineData(3)] // every other division
		public void CanCreateStandingsRequests(int numRounds)
		{
			List<Team> teams = CreateTeams();
			DivisionSheetConfig config = CreateSheetConfig(numRounds);
			PsoDivisionSheetHelper helper = new PsoDivisionSheetHelper(config);
			StandingsTableRequestCreator creator = new StandingsTableRequestCreator(helper, CreateStandingsRequestCreatorFactory(teams, config));

			PoolPlayInfo info = new PoolPlayInfo(teams);
			info = creator.CreateStandingsRequests(config, info, teams, START_ROW_IDX);
			Assert.Equal(NUM_STANDINGS_FORMULAS + teams.Count() + 1, info.UpdateSheetRequests.Count); // there are 14 columns total, but we don't have formulas for 3 of them (yellow/red cards, tiebreaker), then there are head-to-head formulas for each team and a resize request at the end
			IEnumerable<Request> standingsFormulaRequests = info.UpdateSheetRequests.Where(x => x.RepeatCell != null && x.RepeatCell.Range.StartColumnIndex <= helper.GetColumnIndexByHeader(Constants.HDR_GOAL_DIFF));
			Assert.Equal(NUM_STANDINGS_FORMULAS, standingsFormulaRequests.Count());

			// confirm that we have a request for each column that has a formula
			foreach (string hdr in helper.StandingsTableColumns)
			{
				int idx = helper.GetColumnIndexByHeader(hdr);
				Request? rq = standingsFormulaRequests.SingleOrDefault(x => x.RepeatCell.Range.StartColumnIndex == idx);
				switch (hdr)
				{
					case Constants.HDR_YELLOW_CARDS:
					case Constants.HDR_RED_CARDS:
					case Constants.HDR_TIEBREAKER:
						Assert.Null(rq);
						continue;
					default:
						Assert.NotNull(rq);
						break;
				}
			}

			// confirm the columns for the head-to-head tiebreakers
			IEnumerable<Request> headToHeadRequests = info.UpdateSheetRequests.Where(x => x.RepeatCell != null && x.RepeatCell.Range.StartColumnIndex > helper.GetColumnIndexByHeader(Constants.HDR_GOAL_DIFF));
			Assert.Equal(teams.Count, headToHeadRequests.Count());

			// confirm the column indexes for the tiebreaker columns
			int? maxStandingsColIdx = standingsFormulaRequests.Max(x => x.RepeatCell.Range.StartColumnIndex);
			Assert.NotNull(maxStandingsColIdx);
			int? minHeadToHeadColIdx = headToHeadRequests.Min(x => x.RepeatCell.Range.StartColumnIndex);
			Assert.NotNull(minHeadToHeadColIdx);
			Assert.Equal(maxStandingsColIdx!.Value + 1, minHeadToHeadColIdx!.Value);
			int? maxHeadToHeadColIdx = headToHeadRequests.Max(x => x.RepeatCell.Range.StartColumnIndex);
			Assert.NotNull(maxHeadToHeadColIdx);
			Assert.Equal(maxStandingsColIdx!.Value + teams.Count, maxHeadToHeadColIdx!.Value);
		}

		[Fact]
		public void CanCreateScoringRequests()
		{
			// this test validates the requests for creating the scoring section of the sheet; this method is executed once per round of games.
			// there should be 4 update sheet requests (formulas for team names and points) and 1 update values request (winners/pts headers, which get done here and not with the rest of the headers)
			List<Team> teams = CreateTeams();
			DivisionSheetConfig config = CreateSheetConfig();
			PsoDivisionSheetHelper helper = new PsoDivisionSheetHelper(config);
			ScoreInputsRequestCreator creator = new ScoreInputsRequestCreator(teams.First().DivisionName, helper, CreateStandingsRequestCreatorFactory(teams, config));

			PoolPlayInfo info = new PoolPlayInfo(teams);
			int startRowIndex = START_ROW_IDX;
			info = creator.CreateScoringRequests(config, info, teams, 1, ref startRowIndex);

			Assert.Equal(2, info.UpdateValuesRequests.Count);
			GoogleSheetRow roundLabelRow = info.UpdateValuesRequests.First().Rows.Single();
			Assert.Equal(DivisionSheetCreator.HeaderRowColumns.Count, roundLabelRow.Count);
			Assert.Equal("ROUND 1", roundLabelRow.First().StringValue);
			Assert.All(roundLabelRow, cell => Assert.True(cell.GoogleBackgroundColor.GoogleColorEquals(Colors.SubheaderRowColor)));

			Assert.Single(info.UpdateValuesRequests.Last().Rows);
			IEnumerable<string> headers = info.UpdateValuesRequests.Last().Rows.Single().Select(x => x.StringValue);
			headers.Should().BeEquivalentTo(DivisionSheetCreator.WinnerAndPointsColumns);

			Assert.Equal(5, info.UpdateSheetRequests.Count);
			// the first two are data validation requests for the home/away team inputs
			IEnumerable<Request> dpRequests = info.UpdateSheetRequests.Take(2);
			Assert.All(dpRequests, r => Assert.NotNull(r.SetDataValidation));
			Assert.All(dpRequests, r => Assert.Equal(START_ROW_IDX + 2, r.SetDataValidation.Range.StartRowIndex));
			// the last three are the winner and home/away game points formulas
			Assert.All(info.UpdateSheetRequests.Skip(2), r => Assert.NotNull(r.RepeatCell));

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

			List<Team> teams = CreateTeams();
			teams.AddRange(CreateTeams(POOL_B));
			DivisionSheetConfig config = CreateSheetConfig();
			PsoDivisionSheetHelper helper = new PsoDivisionSheetHelper(config);
			WinnerFormattingRequestsCreator creator = new WinnerFormattingRequestsCreator(teams.First().DivisionName, helper);

			Fixture fixture = new Fixture();
			fixture.Customize<PoolPlayInfo>(c => c.FromFactory(() => new PoolPlayInfo(teams))); // https://stackoverflow.com/questions/26149618/autofixture-customizations-provide-constructor-parameter
			PoolPlayInfo poolInfo = fixture.Build<PoolPlayInfo>()
				.Without(x => x.UpdateSheetRequests)
				.Without(x => x.UpdateValuesRequests)
				.With(x => x.StandingsStartAndEndRowNums, fixture.CreateMany<Tuple<int, int>>(2).ToList())
				.Create();
			ChampionshipInfo info = new ChampionshipInfo(poolInfo);

			SheetRequests requests = creator.CreateWinnerFormattingRequests(config, info);

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
		public void PoolPlayCreatorCanSetChampionshipStartRowIndexCorrectly()
		{
			List<Team> teams = CreateTeams();
			teams.AddRange(CreateTeams(POOL_B));
			DivisionSheetConfig config = new DivisionSheetConfig();
			config.SetupForTeams(8);
			PoolPlayInfo info = new PoolPlayInfo(teams);

			Mock<IScoreSheetHeadersRequestCreator> mockHeadersCreator = new Mock<IScoreSheetHeadersRequestCreator>();
			mockHeadersCreator.Setup(x => x.CreateHeaderRequests(It.IsAny<PoolPlayInfo>(), It.IsAny<string>(), It.IsAny<int>(), It.IsAny<IEnumerable<Team>>()))
				.Returns((PoolPlayInfo ppi, string poolName, int rowIdx, IEnumerable<Team> teams) => ppi);

			Mock<IScoreInputsRequestCreator> mockInputsCreator = new Mock<IScoreInputsRequestCreator>();
			mockInputsCreator.Setup(x => x.CreateScoringRequests(It.IsAny<DivisionSheetConfig>(), It.IsAny<PoolPlayInfo>(), It.IsAny<IEnumerable<Team>>(), It.IsAny<int>(), ref It.Ref<int>.IsAny))
				.Returns(new scoreInputReturns((DivisionSheetConfig cfg, PoolPlayInfo ppi, IEnumerable<Team> ts, int rnd, ref int idx) => 
				{
					idx += cfg.GamesPerRound + 2; // +2 accounts the blank row at the end
					return ppi;
				}));

			Mock<IStandingsTableRequestCreator> mockStandingsCreator = new Mock<IStandingsTableRequestCreator>();
			mockStandingsCreator.Setup(x => x.CreateStandingsRequests(It.IsAny<DivisionSheetConfig>(), It.IsAny<PoolPlayInfo>(), It.IsAny<IEnumerable<Team>>(), It.IsAny<int>()))
				.Returns((DivisionSheetConfig cfg, PoolPlayInfo ppi, IEnumerable<Team> ts, int idx) => ppi);

			PoolPlayRequestCreator creator = new PoolPlayRequestCreator(config, mockHeadersCreator.Object, mockInputsCreator.Object, mockStandingsCreator.Object);
			info = creator.CreatePoolPlayRequests(info).Result;

			Assert.Equal(26, info.ChampionshipStartRowIndex); // this number comes from the 2021 score sheet for 10U Boys, as that was an 8-team division
		}

		// https://stackoverflow.com/questions/1068095/assigning-out-ref-parameters-in-moq
		delegate PoolPlayInfo scoreInputReturns(DivisionSheetConfig cfg, PoolPlayInfo ppi, IEnumerable<Team> ts, int rnd, ref int idx);
	}
}
