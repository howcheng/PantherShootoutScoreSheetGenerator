using System.Data;
using Google.Apis.Sheets.v4.Data;
using GoogleSheetsHelper;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services.Tests
{
	public class ShootoutScoreEntryRequestsCreatorTests : SheetGeneratorTests
	{
		// we are going to break this up into multiple tests

		private DivisionSheetConfig? _divisionConfig;
		private List<Team>? _teams;
		private ShootoutSheetHelper? _helper;
		private ShootoutSheetConfig? _shootoutConfig;
		private PsoDivisionSheetHelper? _divSheetHelper;
		private ShootoutScoringFormulaGenerator? _fg;
		private const int START_IDX = 1;
		private const int SCORE_ENTRY_START_IDX = 3;

		// mocks
		private readonly Mock<IStandingsRequestCreator> _mockStandingsReqCreator = new();
		List<StandingsRequestCreatorConfig> _standingsRequestConfigs = new();
		private readonly Mock<IStandingsRequestCreatorFactory> _mockFactory = new();
		private readonly Mock<IShootoutSortedStandingsListRequestCreator> _mockShootoutSortedStandingsCreator = new();
		Tuple<int, int>? _sortedStandingsStartAndEnd;

		private ShootoutScoreEntryRequestsCreator? _creator;

		private void SetupForTests(int numTeams)
		{
			_divisionConfig = DivisionSheetConfigFactory.GetForTeams(numTeams);
			_divisionConfig.DivisionName = nameof(DivisionSheetConfig.DivisionName);
			_teams = CreateTeams(_divisionConfig);
			_helper = new ShootoutSheetHelper(_divisionConfig);
			_shootoutConfig = new ShootoutSheetConfig();
			// this normally would have been set in the ShootoutSheetService
			_shootoutConfig.FirstTeamSheetCells.Add(_divisionConfig.DivisionName, _teams.First().TeamSheetCell);
			_divSheetHelper = new PsoDivisionSheetHelper(_divisionConfig);
			_fg = new ShootoutScoringFormulaGenerator(_helper);

			_mockStandingsReqCreator.Setup(x => x.CreateRequest(It.IsAny<StandingsRequestCreatorConfig>()))
				.Callback((StandingsRequestCreatorConfig cfg) => _standingsRequestConfigs.Add(cfg))
				.Returns((StandingsRequestCreatorConfig cfg) => new Request()
				{
					AddBanding = new(), // this is just so we can identify these later
				});
			_mockFactory.Setup(x => x.GetRequestCreator(It.IsAny<string>())).Returns(_mockStandingsReqCreator.Object);
			_mockShootoutSortedStandingsCreator.Setup(x => x.CreateSortedStandingsListRequest(It.IsAny<Tuple<int, int>>()))
				.Callback((Tuple<int, int> startAndEnd) => _sortedStandingsStartAndEnd = startAndEnd)
				.Returns(new UpdateRequest("asdf")); // dummy value

			_creator = new ShootoutScoreEntryRequestsCreator(_shootoutConfig, _helper, _divSheetHelper
				, _mockFactory.Object
				, _mockShootoutSortedStandingsCreator.Object
				, _fg);
		}

		private PoolPlayInfo CreatePoolPlayInfo(out int startRow)
		{
			PoolPlayInfo poolPlayInfo = new(_teams!);
			startRow = START_IDX + 1; // round header row
			for (int i = 0; i < _divisionConfig!.NumberOfPools; i++)
			{
				List<Tuple<int, int>> startAndEndRows = new();
				poolPlayInfo.ScoreEntryStartAndEndRowNums.Add($"POOL{i}", startAndEndRows);
				for (int j = 0; j < _divisionConfig.NumberOfGameRounds; j++)
				{
					startAndEndRows.Add(new Tuple<int, int>(startRow + 1, startRow + _divisionConfig.GamesPerRound));
					startRow += 4; // including the blank row, the next header row
				}
				startRow += 1; // for the pool row header
			}
			poolPlayInfo.StandingsStartAndEndRowNums.Add(new Tuple<int, int>(START_IDX + 2, START_IDX + _divisionConfig.NumberOfTeams - 1));
			return poolPlayInfo;
		}

		private List<Tuple<int, int>> GetExpectedStartAndEndColumnIndexes()
		{
			List<Tuple<int, int>> expectedStartAndEndColIndexes = new()
			{
				new(8, 11), // I, L
				new(12, 15), // M, P
				new(16, 19), // Q, T
			};
			if (_divisionConfig!.NumberOfShootoutRounds == 4)
				expectedStartAndEndColIndexes.Add(new(20, 23)); // U, X
			return expectedStartAndEndColIndexes;
		}

		[Theory]
		[InlineData(8)]
		[InlineData(10)]
		public void CanCreateRoundHeaderRow(int numTeams)
		{
			SetupForTests(numTeams);
			PoolPlayInfo poolPlayInfo = CreatePoolPlayInfo(out int startRow);
			SheetRequests requests = _creator!.CreateScoreEntryRequests(poolPlayInfo, START_IDX);

			UpdateRequest? headerRequest = requests.UpdateValuesRequests.FirstOrDefault();
			Assert.NotNull(headerRequest);
			Assert.Single(headerRequest!.Rows);
			GoogleSheetRow headerRow = headerRequest.Rows.Single();
			Assert.Equal(4 * _divisionConfig!.NumberOfShootoutRounds, headerRow.Count);
			for (int i = 1; i <= _divisionConfig.NumberOfShootoutRounds; i++)
			{
				Assert.Contains(headerRow, cell => cell.StringValue == $"ROUND {i}");
			}
			Assert.All(headerRow, cell => cell.GoogleBackgroundColor.Should().BeEquivalentTo(Colors.SubheaderRowColor.ToGoogleColor()));
		}

		[Theory]
		[InlineData(8)]
		[InlineData(10)]
		public void CanCreateScoreEntryRequests(int numTeams)
		{
			SetupForTests(numTeams);
			PoolPlayInfo poolPlayInfo = CreatePoolPlayInfo(out int startRow);
			SheetRequests requests = _creator!.CreateScoreEntryRequests(poolPlayInfo, START_IDX);

			int shootoutEntryRequestsPerRound = _divisionConfig!.NumberOfShootoutRounds * _divisionConfig.GamesPerRound * _divisionConfig.NumberOfPools;
			List<Request> scoreEntryRequests = requests.UpdateSheetRequests.Take(shootoutEntryRequestsPerRound).ToList();
			Assert.All(scoreEntryRequests, rq => Assert.NotNull(rq.RepeatCell));
			List<Tuple<int, int>> expectedStartAndEndRows = poolPlayInfo.ScoreEntryStartAndEndRowNums.SelectMany(x => x.Value).ToList();
			List<Tuple<int, int>> expectedStartAndEndColIndexes = GetExpectedStartAndEndColumnIndexes();

			int expectedStartAndEndRowsIdx = 0, expectedStartAndEndColumnsIdx = 0;
			for (int i = 0; i < scoreEntryRequests.Count; i += 2)
			{
				Request homeTeamReq = scoreEntryRequests[i];
				Request awayTeamReq = scoreEntryRequests[i + 1];

				// verify formulas
				Tuple<int, int> expectedStartAndEndRow = expectedStartAndEndRows[expectedStartAndEndRowsIdx];
				string expectedHomeValue = $"='{_divisionConfig.DivisionName}'!A{expectedStartAndEndRow.Item1}";
				Assert.Equal(expectedHomeValue, homeTeamReq.RepeatCell.Cell.UserEnteredValue.FormulaValue);

				string expectedAwayValue = $"='{_divisionConfig.DivisionName}'!D{expectedStartAndEndRow.Item1}";
				Assert.Equal(expectedAwayValue, awayTeamReq.RepeatCell.Cell.UserEnteredValue.FormulaValue);

				// verify ranges
				int expectedStartRowIdx = i < scoreEntryRequests.Count / 2 ? START_IDX + 1 : START_IDX + 3; // +1 for the header row
				Tuple<int, int> expectedStartAndEndCol = expectedStartAndEndColIndexes[expectedStartAndEndColumnsIdx];
				Assert.Equal(expectedStartRowIdx, homeTeamReq.RepeatCell.Range.StartRowIndex);
				Assert.Equal(expectedStartAndEndCol.Item1, homeTeamReq.RepeatCell.Range.StartColumnIndex);
				Assert.Equal(expectedStartRowIdx, awayTeamReq.RepeatCell.Range.StartRowIndex);
				Assert.Equal(expectedStartAndEndCol.Item2 + 1, awayTeamReq.RepeatCell.Range.EndColumnIndex); // +1 because that's how it works

				if (expectedStartAndEndRowsIdx++ == _divisionConfig.NumberOfShootoutRounds - 1 && numTeams == 10)
					expectedStartAndEndRowsIdx += 1; // for 10 teams, there are 5 game rounds, and there's no shootout for the last round
				if (expectedStartAndEndColumnsIdx++ == expectedStartAndEndColIndexes.Count - 1)
					expectedStartAndEndColumnsIdx = 0;
			}
		}

		[Theory]
		[InlineData(8)]
		[InlineData(10)]
		public void CanCreateScoreEntryResizeRequests(int numTeams)
		{
			SetupForTests(numTeams);
			PoolPlayInfo poolPlayInfo = CreatePoolPlayInfo(out int startRow);
			SheetRequests requests = _creator!.CreateScoreEntryRequests(poolPlayInfo, START_IDX);

			List<Request> resizeRequests = requests.UpdateSheetRequests.Where(x => x.UpdateDimensionProperties != null).ToList();
			Assert.Equal(_divisionConfig!.NumberOfShootoutRounds * 2 + ShootoutSheetHelper.TiebreakerColumns.Count, resizeRequests.Count);
			List<Request> scoreEntryResizeRequests = resizeRequests.Take(_divisionConfig.NumberOfShootoutRounds * 2).ToList(); // the last are the tiebreaker columns
			List<Tuple<int, int>> expectedStartAndEndColIndexes = GetExpectedStartAndEndColumnIndexes();

			for (int i = 0; i < scoreEntryResizeRequests.Count; i += 2)
			{
				Request homeTeamReq = scoreEntryResizeRequests[i];
				Request awayTeamReq = scoreEntryResizeRequests[i + 1];
				Tuple<int, int> expectedStartAndEndCol = expectedStartAndEndColIndexes[i / 2];
				Assert.Equal(expectedStartAndEndCol.Item1 + 1, homeTeamReq.UpdateDimensionProperties.Range.StartIndex);
				Assert.Equal(expectedStartAndEndCol.Item2 - 1, awayTeamReq.UpdateDimensionProperties.Range.StartIndex);
			}
		}

		[Theory]
		[InlineData(8)]
		[InlineData(10)]
		public void CanCreateScoreDisplayRequests(int numTeams)
		{
			SetupForTests(numTeams);
			PoolPlayInfo poolPlayInfo = CreatePoolPlayInfo(out int startRow);
			SheetRequests requests = _creator!.CreateScoreEntryRequests(poolPlayInfo, START_IDX);

			List<Request> formulaRequests = requests.UpdateSheetRequests.Where(x => x.RepeatCell != null).ToList();
			// skip the score entry requests
			List<Request> scoreDisplayRequests = formulaRequests.Skip(_divisionConfig!.NumberOfShootoutRounds * 2 * _divisionConfig.NumberOfPools).ToList();
			List<Tuple<int, int>> expectedStartAndEndRows = poolPlayInfo.ScoreEntryStartAndEndRowNums.SelectMany(x => x.Value).ToList();
			List<Tuple<int, int>> expectedScoreDisplayStartAndEndRows = new()
			{
				new(3, 4), // pool 1
				new(5, 6), // pool 2
			};
			List<Tuple<int, int>> expectedStartAndEndColIndexes = GetExpectedStartAndEndColumnIndexes();
			for (int i = 0; i < scoreDisplayRequests.Count; i++)
			{
				// first half of the requests are for the first pool, second half for the second pool
				int poolIdx = i < scoreDisplayRequests.Count / 2 ? 0 : 1;
				Request req = scoreDisplayRequests[i];
				// verify formula
				string formula = req.RepeatCell.Cell.UserEnteredValue.FormulaValue;
				Tuple<int, int> startAndEnd = expectedScoreDisplayStartAndEndRows[poolIdx];
				int expectedStartAndEndRowsIdx = i;
				int roundNum = i + 1;
				if (poolIdx == 1)
				{
					// second pool
					expectedStartAndEndRowsIdx -= _divisionConfig.NumberOfShootoutRounds;
					roundNum -= _divisionConfig.NumberOfShootoutRounds;
				}

				int expectedHomeTeamIdx = expectedStartAndEndColIndexes[expectedStartAndEndRowsIdx].Item1;
				string expectedHomeTeamColName = Utilities.ConvertIndexToColumnName(expectedHomeTeamIdx);
				int expectedStartRowNum = startAndEnd.Item1;
				int expectedEndRowNum = startAndEnd.Item2;
				string expectedHomeTeamCellRange = Utilities.CreateCellRangeString(expectedHomeTeamColName, expectedStartRowNum, expectedEndRowNum, CellRangeOptions.FixRow);
				Assert.Contains(expectedHomeTeamCellRange, formula);

				string expectedHomeScoreColName = Utilities.ConvertIndexToColumnName(expectedHomeTeamIdx + 1);
				string expectedAwayScoreColName = Utilities.ConvertIndexToColumnName(expectedHomeTeamIdx + 2);
				string expectedScoreCellRange = Utilities.CreateCellRangeString(expectedHomeScoreColName, expectedStartRowNum, expectedAwayScoreColName, expectedEndRowNum, CellRangeOptions.FixRow);
				Assert.Contains(expectedScoreCellRange, formula);

				string expectedAwayTeamColName = Utilities.ConvertIndexToColumnName(expectedHomeTeamIdx + 3);
				string expectedAwayTeamCellRange = Utilities.CreateCellRangeString(expectedAwayTeamColName, expectedStartRowNum, expectedEndRowNum, CellRangeOptions.FixRow);
				Assert.Contains(expectedAwayTeamCellRange, formula);

				string expectedTeamSheetCell = poolPlayInfo.Pools!.ElementAt(poolIdx).First().TeamSheetCell;
				Assert.Contains(expectedTeamSheetCell, formula);

				// verify ranges
				int expectedStartRowIdx = poolIdx == 0 ? START_IDX + 1 : START_IDX + _divisionConfig.TeamsPerPool + 1; // +1 for the header row
				Assert.Equal(expectedStartRowIdx, req.RepeatCell.Range.StartRowIndex);
				// expected column idx for the score entry column is the same as the round number
				Assert.Equal(roundNum, req.RepeatCell.Range.StartColumnIndex);
			}
		}

		[Theory]
		[InlineData(8)]
		[InlineData(10)]
		public void CanCreateTiebreakerRequests(int numTeams)
		{
			SetupForTests(numTeams);
			PoolPlayInfo poolPlayInfo = CreatePoolPlayInfo(out int startRow);
			SheetRequests requests = _creator!.CreateScoreEntryRequests(poolPlayInfo, START_IDX);

			List<Request> resizeRequests = requests.UpdateSheetRequests.Where(x => x.UpdateDimensionProperties != null).ToList();
			List<Request> tiebreakerResizeRequests = resizeRequests.Skip(_divisionConfig!.NumberOfShootoutRounds * 2).ToList();
			Assert.Equal(ShootoutSheetHelper.TiebreakerColumns.Count, tiebreakerResizeRequests.Count);

			UpdateRequest? tiebreakerHeadersRequest = requests.UpdateValuesRequests.Skip(1).FirstOrDefault();
			Assert.NotNull(tiebreakerHeadersRequest);
			Assert.Single(tiebreakerHeadersRequest!.Rows);
			Assert.Equal(ShootoutSheetHelper.TiebreakerColumns.Count, tiebreakerHeadersRequest!.Rows.Single().Count);
			Assert.Equal(25, tiebreakerHeadersRequest.ColumnStart);
			Assert.Equal(1, tiebreakerHeadersRequest.RowStart);

			int expectedStandingsRequestsCount = ShootoutSheetHelper.TiebreakerColumns.Count + 1; // +1 for the rank request
			List<Request> standingsRequests = requests.UpdateSheetRequests.Where(x => x.AddBanding != null).ToList(); // we added the AddBanding earlier to identify these
			Assert.Equal(expectedStandingsRequestsCount, standingsRequests.Count);
			Assert.Equal(expectedStandingsRequestsCount, _standingsRequestConfigs.Count);
			Assert.All(_standingsRequestConfigs.Cast<ShootoutTiebreakerRequestCreatorConfig>(), x =>
			{
				Assert.Equal(SCORE_ENTRY_START_IDX, x.StartGamesRowNum);
				Assert.Equal($"A{SCORE_ENTRY_START_IDX}", x.FirstTeamsSheetCell);
			});
		}

		[Theory]
		[InlineData(8)]
		[InlineData(10)]
		public void CanCreateShootoutSortedStandingsList(int numTeams)
		{
			SetupForTests(numTeams);
			PoolPlayInfo poolPlayInfo = CreatePoolPlayInfo(out int startRow);
			SheetRequests requests = _creator!.CreateScoreEntryRequests(poolPlayInfo, START_IDX);

			// verify range for sorted standings list
			Assert.NotNull(_sortedStandingsStartAndEnd);
			Assert.Equal(SCORE_ENTRY_START_IDX, _sortedStandingsStartAndEnd!.Item1);
			Assert.Equal(SCORE_ENTRY_START_IDX + numTeams - 1, _sortedStandingsStartAndEnd!.Item2);
		}

		[Theory]
		[InlineData(8)]
		[InlineData(10)]
		public void CanCreateWinnerFormattingRequests(int numTeams)
		{
			SetupForTests(numTeams);
			PoolPlayInfo poolPlayInfo = CreatePoolPlayInfo(out int startRow);
			SheetRequests requests = _creator!.CreateScoreEntryRequests(poolPlayInfo, START_IDX);

			List<Request> winnerFormatRequests = requests.UpdateSheetRequests.Where(x => x.AddConditionalFormatRule != null).ToList();
			Assert.Equal(2, winnerFormatRequests.Count);

			VerifyWinnerFormattingRequest(winnerFormatRequests.First(), numTeams, false);
			VerifyWinnerFormattingRequest(winnerFormatRequests.Last(), numTeams, true);
		}

		private void VerifyWinnerFormattingRequest(Request request, int numTeams, bool forWinner)
		{
			int startRowIdx = START_IDX + 1;
			ConditionalFormatRule? rule = request.AddConditionalFormatRule!.Rule;
			Assert.NotNull(rule);
			GridRange range = rule.Ranges.Single();
			Assert.Equal(startRowIdx, range.StartRowIndex);
			Assert.Equal(startRowIdx + numTeams, range.EndRowIndex);
			Assert.Equal(0, range.StartColumnIndex);
			Assert.Equal(7, range.EndColumnIndex);

			Assert.Single(rule.BooleanRule.Condition.Values);
			ConditionValue condition = rule.BooleanRule.Condition.Values.Single();
			Assert.Contains("$G3=1", condition.UserEnteredValue); // for the rank
			int startRowNum = startRowIdx + 1;
			int endRowNum = startRowIdx + numTeams;

			string endColName = numTeams == 8 ? "D": "E";
			string scoreRange = $"$B${startRowNum}:${endColName}${endRowNum}";
			string countFunc = forWinner ? "COUNT" : "COUNTBLANK";
			Assert.Contains($"{countFunc}({scoreRange}", condition.UserEnteredValue);

			string teamCountCheck = $"COUNTA($A${startRowNum}:$A${endRowNum})*3";
			Assert.Contains(teamCountCheck, condition.UserEnteredValue);

		}
	}
}
