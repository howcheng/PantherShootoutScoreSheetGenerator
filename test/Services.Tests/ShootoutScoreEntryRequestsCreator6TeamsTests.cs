using System.Data;
using Google.Apis.Sheets.v4.Data;
using GoogleSheetsHelper;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services.Tests
{
	public class ShootoutScoreEntryRequestsCreator6TeamsTests : SheetGeneratorTests
	{
		private const int NUM_TEAMS = 6;
		private const int SHOOTOUT_ROUNDS = 4;
		private const int START_IDX = 1;
		private const int SCORE_ENTRY_START_IDX = 3;

		private DivisionSheetConfig? _divisionConfig;
		private List<Team>? _teams;
		private ShootoutSheetHelper? _helper;
		private ShootoutSheetConfig? _shootoutConfig;
		private PsoDivisionSheetHelper? _divSheetHelper;
		private ShootoutScoringFormulaGenerator? _fg;

		// mocks
		private readonly Mock<IStandingsRequestCreator> _mockStandingsReqCreator = new();
		List<StandingsRequestCreatorConfig> _standingsRequestConfigs = new();
		private readonly Mock<IStandingsRequestCreatorFactory> _mockFactory = new();
		private readonly Mock<IShootoutSortedStandingsListRequestCreator> _mockShootoutSortedStandingsCreator = new();
		Tuple<int, int>? _sortedStandingsStartAndEnd;

		private ShootoutScoreEntryRequestsCreator6Teams? _creator;

		private void SetupForTests()
		{
			_divisionConfig = DivisionSheetConfigFactory.GetForTeams(NUM_TEAMS);
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

			_creator = new ShootoutScoreEntryRequestsCreator6Teams(_shootoutConfig, _helper, _divSheetHelper
				, _mockFactory.Object
				, _mockShootoutSortedStandingsCreator.Object
				, _fg);
		}

		private ChampionshipInfo CreateChampionshipInfo(out int startRow)
		{
			// Create PoolPlayInfo with properly grouped teams
			PoolPlayInfo poolPlayInfo = new(_teams!);
			startRow = START_IDX + 1; // round header row (row 2)
			
			// For 6-team divisions:
			// - 2 pools of 3 teams each
			// - 3 game rounds (pool play)
			// - 1 game per round per pool
			// Each round takes 3 rows: header + game + blank
			for (int i = 0; i < _divisionConfig!.NumberOfPools; i++)
			{
				List<Tuple<int, int>> startAndEndRows = new();
				poolPlayInfo.ScoreEntryStartAndEndRowNums.Add($"POOL{i}", startAndEndRows);
				for (int j = 0; j < _divisionConfig.NumberOfGameRounds; j++)
				{
					// startRow is the header row, so game row is startRow + 1
					startAndEndRows.Add(new Tuple<int, int>(startRow + 1, startRow + 1));
					startRow += 3; // header + game + blank = 3 rows per round
				}
				startRow += 1; // for the separator between pools
			}
			
			// Add standings for both pools
			poolPlayInfo.StandingsStartAndEndRowNums.Add(new Tuple<int, int>(START_IDX + 2, START_IDX + _divisionConfig.TeamsPerPool));
			poolPlayInfo.StandingsStartAndEndRowNums.Add(new Tuple<int, int>(START_IDX + 11, START_IDX + 11 + _divisionConfig.TeamsPerPool - 1));
			
			// Create ChampionshipInfo with semifinal row numbers
			// The ChampionshipInfo constructor takes PoolPlayInfo and copies the Pools property
			ChampionshipInfo champInfo = new ChampionshipInfo(poolPlayInfo);
			
			// Verify Pools was copied
			if (champInfo.Pools == null || !champInfo.Pools.Any())
			{
				throw new InvalidOperationException("ChampionshipInfo.Pools was not properly initialized");
			}
			
			// Championship section starts after pool play (row 21 in screenshot)
			champInfo.ChampionshipStartRowIndex = startRow;
			
			// Based on ChampionshipRequestCreator6Teams structure and screenshot:
			// Row 21: Finals label row
			// Row 22: Consolation subheader
			// Row 23: Consolation game (ConsolationGameRowNum)
			// Row 24: Semifinals subheader
			// Row 25: Semifinal 1 (Semifinal1RowNum)
			// Row 26: Semifinal 2 (Semifinal2RowNum)
			// Row 27: 3rd-place game subheader
			// Row 28: 3rd-place game (ThirdPlaceGameRowNum) - NO SHOOTOUT
			// Row 29: Final subheader
			// Row 30: Final game (ChampionshipGameRowNum) - NO SHOOTOUT
			
			int champStartRowNum = startRow + 1;
			champInfo.ConsolationGameRowNum = champStartRowNum + 2; // after label + consolation subheader
			champInfo.Semifinal1RowNum = champStartRowNum + 4; // after label + consolation subheader + consolation game + semifinals subheader
			champInfo.Semifinal2RowNum = champStartRowNum + 5;
			champInfo.ThirdPlaceGameRowNum = champStartRowNum + 7; // after semifinals + 3rd place subheader
			champInfo.ChampionshipGameRowNum = champStartRowNum + 9; // after 3rd place + final subheader
			
			return champInfo;
		}

		private List<Tuple<int, int>> GetExpectedStartAndEndColumnIndexes()
		{
			// For 4 rounds: ROUND 1, ROUND 2, ROUND 3, ROUND 4
			// Each round has: Home Team, Home Goals, Away Goals, Away Team (4 columns)
			List<Tuple<int, int>> expectedStartAndEndColIndexes = new()
			{
				new(8, 11),   // Round 1: I (8), L (11)
				new(12, 15),  // Round 2: M (12), P (15)
				new(16, 19),  // Round 3: Q (16), T (19)
				new(20, 23),  // Round 4: U (20), X (23)
			};
			return expectedStartAndEndColIndexes;
		}

		[Fact]
		public void CanCreateRoundHeaderRow()
		{
			SetupForTests();
			ChampionshipInfo champInfo = CreateChampionshipInfo(out int startRow);
			SheetRequests requests = _creator!.CreateScoreEntryRequests(champInfo, START_IDX);

			UpdateRequest? headerRequest = requests.UpdateValuesRequests.FirstOrDefault();
			Assert.NotNull(headerRequest);
			Assert.Single(headerRequest!.Rows);
			GoogleSheetRow headerRow = headerRequest.Rows.Single();
			
			// 6-team divisions have 4 rounds, each with 4 cells (home team, home score, away score, away team)
			Assert.Equal(4 * SHOOTOUT_ROUNDS, headerRow.Count);
			
			for (int i = 1; i <= SHOOTOUT_ROUNDS; i++)
			{
				Assert.Contains(headerRow, cell => cell.StringValue == $"ROUND {i}");
			}
			Assert.All(headerRow, cell => cell.GoogleBackgroundColor.Should().BeEquivalentTo(Colors.SubheaderRowColor.ToGoogleColor()));
		}

		[Fact]
		public void CanCreateScoreEntryRequestsForPoolPlayRounds()
		{
			SetupForTests();
			ChampionshipInfo champInfo = CreateChampionshipInfo(out int startRow);
			SheetRequests requests = _creator!.CreateScoreEntryRequests(champInfo, START_IDX);

			// For 6-team divisions:
			// - Rounds 1-3 are pool play: 3 rounds * 1 game per round * 2 pools * 2 requests (home/away) = 12 requests
			// - Round 4 is championship: 3 games (consolation + 2 semifinals) * 2 requests (home/away) = 6 requests
			// Total: 18 score entry requests
			int expectedPoolPlayRequests = _divisionConfig!.NumberOfGameRounds * _divisionConfig.GamesPerRound * _divisionConfig.NumberOfPools * 2; // *2 for home/away
			int expectedRound4Requests = 3 * 2; // 3 games (consolation + 2 semifinals) * 2 for home/away
			int totalExpectedRequests = expectedPoolPlayRequests + expectedRound4Requests;

			List<Request> scoreEntryRequests = requests.UpdateSheetRequests
				.Where(x => x.RepeatCell != null)
				.Take(totalExpectedRequests)
				.ToList();
			
			Assert.Equal(totalExpectedRequests, scoreEntryRequests.Count);

			// The implementation organizes score entry horizontally by pool:
			// - Pool 1: All rounds on row 2 in different columns (Round 1: I-L, Round 2: M-P, Round 3: Q-T)
			// - Pool 2: All rounds on row 3 in different columns (Round 1: I-L, Round 2: M-P, Round 3: Q-T)
			// Division sheet rows for each pool's rounds: Pool 1: [3, 6, 9], Pool 2: [13, 16, 19]
			List<Tuple<int, int>> expectedStartAndEndColIndexes = GetExpectedStartAndEndColumnIndexes();

			// Check first 12 requests (pool play rounds)
			int requestIdx = 0;
			
			// Process each pool
			for (int poolIdx = 0; poolIdx < _divisionConfig.NumberOfPools; poolIdx++)
			{
				int shootoutSheetRowIdx = START_IDX + 1 + poolIdx; // Pool 1 at row 2, Pool 2 at row 3
				
				// Get the division sheet rows for this pool
				List<int> divisionSheetRows = poolIdx == 0 
					? new List<int> { 3, 6, 9 }    // Pool 1 (Pool G) game rows
					: new List<int> { 13, 16, 19 }; // Pool 2 (Pool H) game rows
				
				// For each round in this pool
				for (int roundNum = 1; roundNum <= _divisionConfig.NumberOfGameRounds; roundNum++)
				{
					Request homeTeamReq = scoreEntryRequests[requestIdx];
					Request awayTeamReq = scoreEntryRequests[requestIdx + 1];

					// Verify formulas reference the correct division sheet row
					int expectedDivisionSheetRow = divisionSheetRows[roundNum - 1];
					string expectedHomeValue = $"='{_divisionConfig.DivisionName}'!A{expectedDivisionSheetRow}";
					Assert.Equal(expectedHomeValue, homeTeamReq.RepeatCell.Cell.UserEnteredValue.FormulaValue);

					string expectedAwayValue = $"='{_divisionConfig.DivisionName}'!D{expectedDivisionSheetRow}";
					Assert.Equal(expectedAwayValue, awayTeamReq.RepeatCell.Cell.UserEnteredValue.FormulaValue);

					// Verify ranges on the shootout sheet - all rounds for this pool should be at the same row
					Tuple<int, int> expectedStartAndEndCol = expectedStartAndEndColIndexes[roundNum - 1];
					Assert.Equal(shootoutSheetRowIdx, homeTeamReq.RepeatCell.Range.StartRowIndex);
					Assert.Equal(expectedStartAndEndCol.Item1, homeTeamReq.RepeatCell.Range.StartColumnIndex);
					Assert.Equal(shootoutSheetRowIdx, awayTeamReq.RepeatCell.Range.StartRowIndex);
					Assert.Equal(expectedStartAndEndCol.Item2 + 1, awayTeamReq.RepeatCell.Range.EndColumnIndex); // +1 because that's how it works

					requestIdx += 2; // move to next pair
				}
			}
		}

		[Fact]
		public void CanCreateScoreEntryRequestsForRound4()
		{
			SetupForTests();
			ChampionshipInfo champInfo = CreateChampionshipInfo(out int startRow);
			SheetRequests requests = _creator!.CreateScoreEntryRequests(champInfo, START_IDX);

			// Get Round 4 requests (last 6 requests for score entry)
			int expectedPoolPlayRequests = _divisionConfig!.NumberOfGameRounds * _divisionConfig.GamesPerRound * _divisionConfig.NumberOfPools * 2;
			List<Request> round4Requests = requests.UpdateSheetRequests
				.Where(x => x.RepeatCell != null)
				.Skip(expectedPoolPlayRequests)
				.Take(6)
				.ToList();

			Assert.Equal(6, round4Requests.Count);

			// Round 4 column indexes (column U-X, which is 20-23)
			Tuple<int, int> round4Cols = GetExpectedStartAndEndColumnIndexes()[3];

			// Round 4 games start at the same row as pool play (row 2) for horizontal alignment
			// Consolation game at row index 2
			Request consolationHomeReq = round4Requests[0];
			Request consolationAwayReq = round4Requests[1];
			string expectedConsolationHome = $"='{_divisionConfig.DivisionName}'!A{champInfo.ConsolationGameRowNum}";
			string expectedConsolationAway = $"='{_divisionConfig.DivisionName}'!D{champInfo.ConsolationGameRowNum}";
			Assert.Equal(expectedConsolationHome, consolationHomeReq.RepeatCell.Cell.UserEnteredValue.FormulaValue);
			Assert.Equal(expectedConsolationAway, consolationAwayReq.RepeatCell.Cell.UserEnteredValue.FormulaValue);
			Assert.Equal(START_IDX + 1, consolationHomeReq.RepeatCell.Range.StartRowIndex); // Row index 2 - horizontally aligned with pool play
			Assert.Equal(round4Cols.Item1, consolationHomeReq.RepeatCell.Range.StartColumnIndex);

			// Semifinal 1 at row index 3 (next row after consolation)
			Request semi1HomeReq = round4Requests[2];
			Request semi1AwayReq = round4Requests[3];
			string expectedSemi1Home = $"='{_divisionConfig.DivisionName}'!A{champInfo.Semifinal1RowNum}";
			string expectedSemi1Away = $"='{_divisionConfig.DivisionName}'!D{champInfo.Semifinal1RowNum}";
			Assert.Equal(expectedSemi1Home, semi1HomeReq.RepeatCell.Cell.UserEnteredValue.FormulaValue);
			Assert.Equal(expectedSemi1Away, semi1AwayReq.RepeatCell.Cell.UserEnteredValue.FormulaValue);
			Assert.Equal(START_IDX + 2, semi1HomeReq.RepeatCell.Range.StartRowIndex); // Row index 3
			Assert.Equal(round4Cols.Item1, semi1HomeReq.RepeatCell.Range.StartColumnIndex);

			// Semifinal 2 at row index 4 (next row after semifinal 1)
			Request semi2HomeReq = round4Requests[4];
			Request semi2AwayReq = round4Requests[5];
			string expectedSemi2Home = $"='{_divisionConfig.DivisionName}'!A{champInfo.Semifinal2RowNum}";
			string expectedSemi2Away = $"='{_divisionConfig.DivisionName}'!D{champInfo.Semifinal2RowNum}";
			Assert.Equal(expectedSemi2Home, semi2HomeReq.RepeatCell.Cell.UserEnteredValue.FormulaValue);
			Assert.Equal(expectedSemi2Away, semi2AwayReq.RepeatCell.Cell.UserEnteredValue.FormulaValue);
			Assert.Equal(START_IDX + 3, semi2HomeReq.RepeatCell.Range.StartRowIndex); // Row index 4
			Assert.Equal(round4Cols.Item1, semi2HomeReq.RepeatCell.Range.StartColumnIndex);
		}
		[Fact]
		public void CanCreateScoreEntryResizeRequests()
		{
			SetupForTests();
			ChampionshipInfo champInfo = CreateChampionshipInfo(out int startRow);
			SheetRequests requests = _creator!.CreateScoreEntryRequests(champInfo, START_IDX);

			List<Request> resizeRequests = requests.UpdateSheetRequests.Where(x => x.UpdateDimensionProperties != null).ToList();
			// 4 rounds * 2 (home/away scores) + tiebreaker columns
			Assert.Equal(SHOOTOUT_ROUNDS * 2 + ShootoutSheetHelper.TiebreakerColumns.Count, resizeRequests.Count);
			
			List<Request> scoreEntryResizeRequests = resizeRequests.Take(SHOOTOUT_ROUNDS * 2).ToList();
			List<Tuple<int, int>> expectedStartAndEndColIndexes = GetExpectedStartAndEndColumnIndexes();

			for (int i = 0; i < scoreEntryResizeRequests.Count; i += 2)
			{
				Request homeScoreReq = scoreEntryResizeRequests[i];
				Request awayScoreReq = scoreEntryResizeRequests[i + 1];
				Tuple<int, int> expectedStartAndEndCol = expectedStartAndEndColIndexes[i / 2];
				
				// Home goals is column index +1 from home team
				Assert.Equal(expectedStartAndEndCol.Item1 + 1, homeScoreReq.UpdateDimensionProperties.Range.StartIndex);
				// Away goals is column index -1 from away team
				Assert.Equal(expectedStartAndEndCol.Item2 - 1, awayScoreReq.UpdateDimensionProperties.Range.StartIndex);
			}
		}

		[Fact]
		public void CanCreateScoreDisplayRequests()
		{
			SetupForTests();
			ChampionshipInfo champInfo = CreateChampionshipInfo(out int startRow);
			SheetRequests requests = _creator!.CreateScoreEntryRequests(champInfo, START_IDX);

			List<Request> formulaRequests = requests.UpdateSheetRequests.Where(x => x.RepeatCell != null).ToList();
			
			// Skip the score entry requests (12 pool play + 6 round 4 = 18 total)
			int totalScoreEntryRequests = (_divisionConfig!.NumberOfGameRounds * _divisionConfig.GamesPerRound * _divisionConfig.NumberOfPools * 2) + 6;
			List<Request> scoreDisplayRequests = formulaRequests.Skip(totalScoreEntryRequests).ToList();
			
			// For 6-team divisions: 2 pools * 4 rounds = 8 score display requests
			Assert.Equal(_divisionConfig.NumberOfPools * SHOOTOUT_ROUNDS, scoreDisplayRequests.Count);

			List<Tuple<int, int>> expectedStartAndEndColIndexes = GetExpectedStartAndEndColumnIndexes();
			
			for (int i = 0; i < scoreDisplayRequests.Count; i++)
			{
				int poolIdx = i < SHOOTOUT_ROUNDS ? 0 : 1;
				Request req = scoreDisplayRequests[i];
				
				// verify formula contains expected elements
				string formula = req.RepeatCell.Cell.UserEnteredValue.FormulaValue;
				
				int roundNum = (i % SHOOTOUT_ROUNDS) + 1;
				Tuple<int, int> expectedColIndexes = expectedStartAndEndColIndexes[roundNum - 1];
				
				string expectedHomeTeamColName = Utilities.ConvertIndexToColumnName(expectedColIndexes.Item1);
				string expectedAwayTeamColName = Utilities.ConvertIndexToColumnName(expectedColIndexes.Item2);
				
				// Formula should contain references to the score entry columns for this round
				Assert.Contains(expectedHomeTeamColName, formula);
				Assert.Contains(expectedAwayTeamColName, formula);
				
				string expectedTeamSheetCell = champInfo.Pools!.ElementAt(poolIdx).First().TeamSheetCell;
				Assert.Contains(expectedTeamSheetCell, formula);

				// verify ranges
				int expectedStartRowIdx = poolIdx == 0 ? START_IDX + 1 : START_IDX + 1 + _divisionConfig.TeamsPerPool;
				Assert.Equal(expectedStartRowIdx, req.RepeatCell.Range.StartRowIndex);
				Assert.Equal(roundNum, req.RepeatCell.Range.StartColumnIndex); // Score display column is same as round number
			}
		}

		[Fact]
		public void CanCreateTiebreakerRequests()
		{
			SetupForTests();
			ChampionshipInfo champInfo = CreateChampionshipInfo(out int startRow);
			SheetRequests requests = _creator!.CreateScoreEntryRequests(champInfo, START_IDX);

			List<Request> resizeRequests = requests.UpdateSheetRequests.Where(x => x.UpdateDimensionProperties != null).ToList();
			List<Request> tiebreakerResizeRequests = resizeRequests.Skip(SHOOTOUT_ROUNDS * 2).ToList();
			Assert.Equal(ShootoutSheetHelper.TiebreakerColumns.Count, tiebreakerResizeRequests.Count);

			UpdateRequest? tiebreakerHeadersRequest = requests.UpdateValuesRequests.Skip(1).FirstOrDefault();
			Assert.NotNull(tiebreakerHeadersRequest);
			Assert.Single(tiebreakerHeadersRequest!.Rows);
			Assert.Equal(ShootoutSheetHelper.TiebreakerColumns.Count, tiebreakerHeadersRequest!.Rows.Single().Count);
			Assert.Equal(25, tiebreakerHeadersRequest.ColumnStart); // Column Y (index 24) for tiebreaker team name
			Assert.Equal(START_IDX, tiebreakerHeadersRequest.RowStart);

			// Now we create separate tiebreaker requests for each pool, so we have:
			// (TiebreakerColumns.Count * NumberOfPools) + 1 rank request
			int expectedStandingsRequestsCount = (ShootoutSheetHelper.TiebreakerColumns.Count * _divisionConfig!.NumberOfPools) + 1;
			List<Request> standingsRequests = requests.UpdateSheetRequests.Where(x => x.AddBanding != null).ToList();
			Assert.Equal(expectedStandingsRequestsCount, standingsRequests.Count);
			Assert.Equal(expectedStandingsRequestsCount, _standingsRequestConfigs.Count);
			
			// Verify each pool gets its own set of tiebreaker requests with correct row numbers
			int expectedPoolStartRowNum = SCORE_ENTRY_START_IDX;
			for (int poolIdx = 0; poolIdx < _divisionConfig.NumberOfPools; poolIdx++)
			{
				int poolStartRowNum = expectedPoolStartRowNum + (poolIdx * _divisionConfig.TeamsPerPool);
				
				// Get configs for this pool (3 tiebreaker columns per pool)
				var poolConfigs = _standingsRequestConfigs
					.Cast<ShootoutTiebreakerRequestCreatorConfig>()
					.Skip(poolIdx * ShootoutSheetHelper.TiebreakerColumns.Count)
					.Take(ShootoutSheetHelper.TiebreakerColumns.Count)
					.ToList();
					
				Assert.All(poolConfigs, x =>
				{
					Assert.Equal(poolStartRowNum, x.StartGamesRowNum);
					Assert.Equal($"A{poolStartRowNum}", x.FirstTeamsSheetCell);
					Assert.Equal(_divisionConfig.TeamsPerPool, x.RowCount);
				});
			}
		}

		[Fact]
		public void CanCreateShootoutSortedStandingsList()
		{
			SetupForTests();
			ChampionshipInfo champInfo = CreateChampionshipInfo(out int startRow);
			SheetRequests requests = _creator!.CreateScoreEntryRequests(champInfo, START_IDX);

			// verify range for sorted standings list
			Assert.NotNull(_sortedStandingsStartAndEnd);
			Assert.Equal(SCORE_ENTRY_START_IDX, _sortedStandingsStartAndEnd!.Item1);
			Assert.Equal(SCORE_ENTRY_START_IDX + NUM_TEAMS - 1, _sortedStandingsStartAndEnd!.Item2);
		}

		[Fact]
		public void CanCreateWinnerFormattingRequests()
		{
			SetupForTests();
			ChampionshipInfo champInfo = CreateChampionshipInfo(out int startRow);
			SheetRequests requests = _creator!.CreateScoreEntryRequests(champInfo, START_IDX);

			List<Request> winnerFormatRequests = requests.UpdateSheetRequests.Where(x => x.AddConditionalFormatRule != null).ToList();
			Assert.Equal(2, winnerFormatRequests.Count);

			VerifyWinnerFormattingRequest(winnerFormatRequests.First(), false);
			VerifyWinnerFormattingRequest(winnerFormatRequests.Last(), true);
		}

		private void VerifyWinnerFormattingRequest(Request request, bool forWinner)
		{
			int startRowIdx = START_IDX + 1;
			ConditionalFormatRule? rule = request.AddConditionalFormatRule!.Rule;
			Assert.NotNull(rule);
			GridRange range = rule.Ranges.Single();
			Assert.Equal(startRowIdx, range.StartRowIndex);
			Assert.Equal(startRowIdx + NUM_TEAMS, range.EndRowIndex);
			Assert.Equal(0, range.StartColumnIndex);
			Assert.Equal(7, range.EndColumnIndex); // Through Rank column (G)

			Assert.Single(rule.BooleanRule.Condition.Values);
			ConditionValue condition = rule.BooleanRule.Condition.Values.Single();
			Assert.Contains("$G3=1", condition.UserEnteredValue); // for the rank
			
			int startRowNum = startRowIdx + 1;
			int endRowNum = startRowIdx + NUM_TEAMS;

			// For 6-team divisions with 4 rounds, the range is $B$3:$E$8
			string scoreRange = $"$B${startRowNum}:$E${endRowNum}";
			string countFunc = forWinner ? "COUNT" : "COUNTBLANK";
			Assert.Contains($"{countFunc}({scoreRange}", condition.UserEnteredValue);

			string teamCountCheck = $"COUNTA($A${startRowNum}:$A${endRowNum})*3";
			Assert.Contains(teamCountCheck, condition.UserEnteredValue);
		}

		[Fact]
		public void ThrowsExceptionIfPoolPlayInfoIsNotChampionshipInfo()
		{
			SetupForTests();
			PoolPlayInfo poolPlayInfo = new(_teams!);
			
			// Should throw ArgumentException when PoolPlayInfo is not a ChampionshipInfo
			Assert.Throws<ArgumentException>(() => _creator!.CreateScoreEntryRequests(poolPlayInfo, START_IDX));
		}
	}
}
