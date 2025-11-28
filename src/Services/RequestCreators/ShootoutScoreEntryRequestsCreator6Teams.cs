using Google.Apis.Sheets.v4.Data;
using GoogleSheetsHelper;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	/// <summary>
	/// Creates the requests to build the scoring entry section of the Shootout sheet for 6-team divisions (score entry, tiebreakers, and sorted standings)
	/// </summary>
	/// <remarks>
	/// For 6-team divisions:
	/// - Round 1-3: Pool play games (2 games per round, all 6 teams participate)
	/// - Round 4: Consolation + Semifinals (3 games, all 6 teams participate)
	/// This ensures each team has exactly 3 shootout opportunities.
	/// Need one instance per division.
	/// </remarks>
	public class ShootoutScoreEntryRequestsCreator6Teams : IShootoutScoreEntryRequestsCreator
	{
		private readonly ShootoutSheetConfig _shootoutSheetConfig;
		private readonly ShootoutSheetHelper _helper;
		private readonly PsoDivisionSheetHelper _divisionSheetHelper;
		private readonly string _divisionName;
		private readonly IStandingsRequestCreatorFactory _requestCreatorFactory;
		private readonly IShootoutSortedStandingsListRequestCreator _sortedStandingsCreator;
		private readonly ShootoutScoringFormulaGenerator _formulaGenerator;
		private readonly DivisionSheetConfig _divisionConfig;
		private const int SHOOTOUT_ROUNDS = 4;

		public ShootoutScoreEntryRequestsCreator6Teams(ShootoutSheetConfig shootoutSheetConfig, StandingsSheetHelper helper, PsoDivisionSheetHelper divSheetHelper
			, IStandingsRequestCreatorFactory requestCreatorFactory, IShootoutSortedStandingsListRequestCreator sortedStandingsCreator
			, FormulaGenerator formulaGenerator)
		{
			_shootoutSheetConfig = shootoutSheetConfig;
			_helper = (ShootoutSheetHelper)helper;
			_divisionSheetHelper = divSheetHelper;
			_divisionName = _helper.DivisionSheetConfig.DivisionName;
			_requestCreatorFactory = requestCreatorFactory;
			_sortedStandingsCreator = sortedStandingsCreator;
			_formulaGenerator = (ShootoutScoringFormulaGenerator)formulaGenerator;
			_divisionConfig = _divisionSheetHelper.DivisionSheetConfig;
		}

		public SheetRequests CreateScoreEntryRequests(PoolPlayInfo info, int startRowIndex)
		{
			SheetRequests ret = new();
			
			// Cast to get access to the semifinal row numbers
			ChampionshipInfo champInfo = info as ChampionshipInfo ?? throw new ArgumentException("PoolPlayInfo must be a ChampionshipInfo for 6-team divisions", nameof(info));

			// header row
			CreateScoreEntryHeaderRow(startRowIndex, ret);
			int scoreEntryStartRowIdx = startRowIndex + 1;

			// score entry rows for all 4 rounds
			List<Tuple<int, int>> shootoutEntryStartAndEndRows;
			CreateScoreEntryRequests(champInfo, ret, scoreEntryStartRowIdx, out shootoutEntryStartAndEndRows);

			// resize columns for all 4 rounds
			for (int i = 0; i < SHOOTOUT_ROUNDS; i++)
			{
				int roundNum = i + 1;
				CreateScoreEntryResizeRequests(ret, roundNum);
			}

			// score display columns
			CreateScoreDisplayRequests(ret, champInfo, shootoutEntryStartAndEndRows);

			// tiebreaker columns
			CreateTiebreakerColumns(startRowIndex, champInfo, ret, shootoutEntryStartAndEndRows);

			// sorted standings list
			Tuple<int, int> shootoutStartAndEnd = new(scoreEntryStartRowIdx + 1, scoreEntryStartRowIdx + _divisionConfig.NumberOfTeams);
			CreateSortedStandingsList(ret, shootoutStartAndEnd);

			// rank column
			CreateRankRequest(scoreEntryStartRowIdx, shootoutStartAndEnd, ret);

			// winner formatting
			CreateWinnerFormattingRequests(false, shootoutStartAndEnd, ret);
			CreateWinnerFormattingRequests(true, shootoutStartAndEnd, ret);

			return ret;
		}

		private void CreateScoreEntryHeaderRow(int startRowIndex, SheetRequests sheetRequests)
		{
			UpdateRequest headerRequest = new(ShootoutConstants.SHOOTOUT_SHEET_NAME)
			{
				RowStart = startRowIndex,
				ColumnStart = _helper.GetColumnIndexForScoreEntry(1),
			};
			GoogleSheetRow headerRow = new();
			for (int i = 0; i < SHOOTOUT_ROUNDS; i++)
			{
				headerRow.Add(new($"ROUND {i + 1}")); // home team
				headerRow.Add(new(string.Empty)); // home score
				headerRow.Add(new(string.Empty)); // away score
				headerRow.Add(new(string.Empty)); // away team
			}
			foreach (GoogleSheetCell cell in headerRow)
			{
				cell.SetSubheaderCellFormatting();
			}
			headerRequest.Rows.Add(headerRow);
			sheetRequests.UpdateValuesRequests.Add(headerRequest);
		}

		private void CreateScoreEntryRequests(ChampionshipInfo info, SheetRequests ret, int scoreEntryStartRowIdx, out List<Tuple<int, int>> shootoutEntryStartAndEndRows)
		{
			shootoutEntryStartAndEndRows = new(SHOOTOUT_ROUNDS);
			int idxCounter = scoreEntryStartRowIdx;
			
			// For 6-team divisions with 2 pools:
			// - Pool 1 at row index scoreEntryStartRowIdx (all rounds in different columns)
			// - Pool 2 at row index scoreEntryStartRowIdx + 1 (all rounds in different columns)
			// - Round 4 championship games also start at scoreEntryStartRowIdx (in Round 4 columns)
			
			// Process each pool's games - similar to standard implementation
			// For 6-team divisions, we create all rounds for a pool on the same row (different columns)
			foreach (List<Tuple<int, int>> scoreEntryStartAndEnd in info.ScoreEntryStartAndEndRowNums.Select(p => p.Value))
			{
				// For each round in this pool, create score entry formulas at the current row
				for (int i = 0; i < _divisionConfig.NumberOfGameRounds; i++)
				{
					Tuple<int, int> scoreStartAndEnd = scoreEntryStartAndEnd[i];
					int roundNum = i + 1;
					CreateScoreEntryRowsForRound(idxCounter, roundNum, ret, scoreStartAndEnd);
				}
				
				// Move to next pool's row
				idxCounter += _divisionConfig.GamesPerRound;
			}
			
			// For rounds 1-3, all pools' games are in the same row range
			int poolPlayEndRowIdx = idxCounter - 1;
			for (int i = 0; i < _divisionConfig.NumberOfGameRounds; i++)
			{
				shootoutEntryStartAndEndRows.Add(new Tuple<int, int>(scoreEntryStartRowIdx + 1, poolPlayEndRowIdx + 1)); // +1 to convert from index to row number
			}
			
			// Round 4: Championship games (consolation + semifinals)
			// These should start at the same row as pool play to maintain horizontal alignment
			CreateRound4ScoreEntries(info, ret, scoreEntryStartRowIdx, out Tuple<int, int> round4StartAndEnd);
			shootoutEntryStartAndEndRows.Add(round4StartAndEnd);
		}

		private void CreateRound4ScoreEntries(ChampionshipInfo info, SheetRequests ret, int startRowIndex, out Tuple<int, int> round4StartAndEnd)
		{
			// Round 4 consists of:
			// 1. Consolation game (ConsolationGameRowNum)
			// 2. Semifinal 1 (Semifinal1RowNum)
			// 3. Semifinal 2 (Semifinal2RowNum)

			int round4StartRowNum = startRowIndex + 1;
			
			// Consolation game
			CreateScoreEntryRowForChampionshipGame(startRowIndex, 4, ret, info.ConsolationGameRowNum);
			startRowIndex++;

			// Semifinal 1
			CreateScoreEntryRowForChampionshipGame(startRowIndex, 4, ret, info.Semifinal1RowNum);
			startRowIndex++;

			// Semifinal 2
			CreateScoreEntryRowForChampionshipGame(startRowIndex, 4, ret, info.Semifinal2RowNum);

			// Round 4 has 3 games
			round4StartAndEnd = new Tuple<int, int>(round4StartRowNum, round4StartRowNum + 2);
		}

		private void CreateScoreEntryRowForChampionshipGame(int startRowIndex, int roundNum, SheetRequests sheetRequests, int gameRowNum)
		{
			// Reference the championship game row on the division sheet
			string homeTeamNameCell = Utilities.CreateCellReference(_divisionSheetHelper.HomeTeamColumnName, gameRowNum, _divisionName);
			string homeTeamFormula = $"={homeTeamNameCell}";
			string awayTeamNameCell = Utilities.CreateCellReference(_divisionSheetHelper.AwayTeamColumnName, gameRowNum, _divisionName);
			string awayTeamFormula = $"={awayTeamNameCell}";

			Request homeTeamReq = RequestCreator.CreateRepeatedSheetFormulaRequest(_shootoutSheetConfig.SheetId, startRowIndex
				, _helper.GetColumnIndexByHeader(Constants.HDR_HOME_TEAM, roundNum), 1, homeTeamFormula);

			Request awayTeamReq = RequestCreator.CreateRepeatedSheetFormulaRequest(_shootoutSheetConfig.SheetId, startRowIndex
				, _helper.GetColumnIndexByHeader(Constants.HDR_AWAY_TEAM, roundNum), 1, awayTeamFormula);

			sheetRequests.UpdateSheetRequests.Add(homeTeamReq);
			sheetRequests.UpdateSheetRequests.Add(awayTeamReq);
		}

		private Tuple<int, int> CreateScoreEntryRowsForRound(int startRowIndex, int roundNum, SheetRequests sheetRequests, Tuple<int, int> scoreEntryStartAndEnd)
		{
			// For 6-team divisions, GamesPerRound = 1, so we create a single formula at the specified row
			int scoreEntryRowNum = scoreEntryStartAndEnd.Item1;
			string homeTeamNameCell = Utilities.CreateCellReference(_divisionSheetHelper.HomeTeamColumnName, scoreEntryRowNum, _divisionName);
			string homeTeamFormula = $"={homeTeamNameCell}";
			string awayTeamNameCell = Utilities.CreateCellReference(_divisionSheetHelper.AwayTeamColumnName, scoreEntryRowNum, _divisionName);
			string awayTeamFormula = $"={awayTeamNameCell}";

			// Create single-cell requests instead of RepeatCellRequest
			Request homeTeamReq = RequestCreator.CreateRepeatedSheetFormulaRequest(_shootoutSheetConfig.SheetId, startRowIndex
				, _helper.GetColumnIndexByHeader(Constants.HDR_HOME_TEAM, roundNum), 1, homeTeamFormula);

			Request awayTeamReq = RequestCreator.CreateRepeatedSheetFormulaRequest(_shootoutSheetConfig.SheetId, startRowIndex
				, _helper.GetColumnIndexByHeader(Constants.HDR_AWAY_TEAM, roundNum), 1, awayTeamFormula);

			sheetRequests.UpdateSheetRequests.Add(homeTeamReq);
			sheetRequests.UpdateSheetRequests.Add(awayTeamReq);

			// return the start and end row of the score entry rows for this round
			return new Tuple<int, int>(startRowIndex + 1, startRowIndex + _divisionConfig.GamesPerRound);
		}

		private void CreateScoreEntryResizeRequests(SheetRequests sheetRequests, int roundNum)
		{
			Request homeScoreResizeReq = RequestCreator.CreateCellWidthRequest(_shootoutSheetConfig.SheetId, Constants.WIDTH_NUM_COL
				, _helper.GetColumnIndexByHeader(Constants.HDR_HOME_GOALS, roundNum));
			Request awayScoreResizeReq = RequestCreator.CreateCellWidthRequest(_shootoutSheetConfig.SheetId, Constants.WIDTH_NUM_COL
				, _helper.GetColumnIndexByHeader(Constants.HDR_AWAY_GOALS, roundNum));
			sheetRequests.UpdateSheetRequests.Add(homeScoreResizeReq);
			sheetRequests.UpdateSheetRequests.Add(awayScoreResizeReq);
		}

		private void CreateScoreDisplayRequests(SheetRequests ret, ChampionshipInfo info, List<Tuple<int, int>> shootoutEntryStartAndEndRows)
		{
			int idxCounter = shootoutEntryStartAndEndRows[0].Item1 - 1; // -1 because the tuple has row numbers
			
			// For 6-team divisions, we have 2 pools of 3 teams each
			for (int i = 0; i < _divisionConfig.NumberOfPools; i++)
			{
				var poolTeams = info.Pools!.ElementAt(i);
				string firstTeamSheetCell = poolTeams.First().TeamSheetCell;
				
				// Rounds 1-4
				for (int j = 0; j < SHOOTOUT_ROUNDS; j++)
				{
					int roundNum = j + 1;
					CreateScoreDisplayColumnsForRound(firstTeamSheetCell, idxCounter, roundNum, ret, shootoutEntryStartAndEndRows[j]);
				}
				idxCounter += _divisionConfig.TeamsPerPool;
			}
		}

		private void CreateScoreDisplayColumnsForRound(string firstTeamSheetCell, int startRowIndex, int roundNum, SheetRequests sheetRequests, Tuple<int, int> scoreEntryStartAndEnd)
		{
			int startRowNum = scoreEntryStartAndEnd.Item1;
			int endRowNum = scoreEntryStartAndEnd.Item2;
			int colIdx = _helper.GetColumnIndexByHeader($"R{roundNum}");
			
			// For 6-team divisions (3 teams per pool), not all teams play in every round (similar to 5-team divisions)
			// Use the 5-teams formula which checks if the team actually played in that round
			string formula = _formulaGenerator.GetScoreDisplayFormulaWithBye(firstTeamSheetCell, startRowNum, endRowNum, roundNum);

			Request req = RequestCreator.CreateRepeatedSheetFormulaRequest(_shootoutSheetConfig.SheetId, startRowIndex, colIdx, _divisionConfig.TeamsPerPool, formula);
			sheetRequests.UpdateSheetRequests.Add(req);
		}

		private void CreateTiebreakerColumns(int startRowIndex, ChampionshipInfo info, SheetRequests sheetRequests, List<Tuple<int, int>> scoreEntryStartAndEndRows)
		{
			// header columns
			UpdateRequest headerRequest = new(ShootoutConstants.SHOOTOUT_SHEET_NAME)
			{
				RowStart = startRowIndex,
				ColumnStart = _helper.GetColumnIndexByHeader(ShootoutConstants.HDR_TIEBREAKER_TEAM_NAME),
			};
			headerRequest.Rows.Add(new GoogleSheetRow
			{
				new GoogleSheetCell(ShootoutConstants.HDR_TIEBREAKER_TEAM_NAME),
				new GoogleSheetCell(Constants.HDR_TIEBREAKER_GOALS_AGAINST),
				new GoogleSheetCell(Constants.HDR_TIEBREAKER_KFTM_WINNER),
			});
			sheetRequests.UpdateValuesRequests.Add(headerRequest);

			// tiebreaker columns
			foreach (string hdr in ShootoutSheetHelper.TiebreakerColumns)
			{
				IStandingsRequestCreator creator = _requestCreatorFactory.GetRequestCreator(hdr);
				// The tiebreaker section starts at the row after the header (startRowIndex + 1 for the index)
				// but formulas use row numbers (1-indexed), so we need startRowIndex + 2
				int startGamesRowNum = startRowIndex + 2;
				ShootoutTiebreakerRequestCreatorConfig config = new()
				{
					SheetId = _shootoutSheetConfig.SheetId,
					SheetStartRowIndex = startRowIndex + 1,
					RowCount = _divisionConfig.NumberOfTeams,
					StartGamesRowNum = startGamesRowNum,
					FirstTeamsSheetCell = Utilities.CreateCellReference(_helper.TeamNameColumnName, startGamesRowNum),
					ScoreEntryStartAndEndRowNums = scoreEntryStartAndEndRows
				};
				Request request = creator.CreateRequest(config);
				sheetRequests.UpdateSheetRequests.Add(request);

				// resize the column
				Request resizeRequest = RequestCreator.CreateCellWidthRequest(_shootoutSheetConfig.SheetId, Constants.WIDTH_WIDE_NUM_COL, _helper.GetColumnIndexByHeader(hdr));
				sheetRequests.UpdateSheetRequests.Add(resizeRequest);
			}
		}

		private void CreateSortedStandingsList(SheetRequests ret, Tuple<int, int> shootoutStartAndEnd)
		{
			UpdateRequest sortedStandingRequest = _sortedStandingsCreator.CreateSortedStandingsListRequest(shootoutStartAndEnd);
			ret.UpdateValuesRequests.Add(sortedStandingRequest);
		}

		private void CreateRankRequest(int startRowIndex, Tuple<int, int> shootoutStartAndEnd, SheetRequests sheetRequests)
		{
			IStandingsRequestCreator creator = _requestCreatorFactory.GetRequestCreator(Constants.HDR_RANK);
			ShootoutTiebreakerRequestCreatorConfig config = new()
			{
				SheetId = _shootoutSheetConfig.SheetId,
				SheetStartRowIndex = startRowIndex,
				RowCount = _divisionConfig.NumberOfTeams,
				StartGamesRowNum = shootoutStartAndEnd.Item1,
				EndGamesRowNum = shootoutStartAndEnd.Item2,
				FirstTeamsSheetCell = Utilities.CreateCellReference(_helper.TeamNameColumnName, shootoutStartAndEnd.Item1),
			};
			Request request = creator.CreateRequest(config);
			sheetRequests.UpdateSheetRequests.Add(request);
		}

		private void CreateWinnerFormattingRequests(bool isForWinner, Tuple<int, int> shootoutStartAndEnd, SheetRequests sheetRequests)
		{
			int startRowNum = shootoutStartAndEnd.Item1;
			int endRowNum = shootoutStartAndEnd.Item2;
			Request request = new Request
			{
				AddConditionalFormatRule = new AddConditionalFormatRuleRequest
				{
					Rule = new ConditionalFormatRule
					{
						Ranges = new List<GridRange>
						{
							new GridRange
							{
								SheetId = _shootoutSheetConfig.SheetId,
								StartRowIndex = startRowNum - 1,
								StartColumnIndex = _helper.GetColumnIndexByHeader(Constants.HDR_TEAM_NAME),
								EndRowIndex = endRowNum,
								EndColumnIndex = _helper.GetColumnIndexByHeader(Constants.HDR_RANK) + 1,
							}
						},
						BooleanRule = new BooleanRule
						{
							Condition = new BooleanCondition
							{
								Type = "CUSTOM_FORMULA",
								Values = new List<ConditionValue>(),
							},
							Format = new CellFormat
							{
								BackgroundColor = Colors.FirstPlaceColor,
							},
						},
					},
				}
			};
			BooleanRule rule = request.AddConditionalFormatRule.Rule.BooleanRule;
			string rankCell = $"${_helper.GetColumnNameByHeader(Constants.HDR_RANK)}{startRowNum}";
			string firstRoundCol = _helper.GetColumnNameByHeader(ShootoutConstants.HDR_ROUND1);
			// For 6-team divisions, we always use 4 rounds
			string lastRoundCol = _helper.GetColumnNameByHeader(ShootoutConstants.HDR_ROUND4);
			string scoreRange = Utilities.CreateCellRangeString(firstRoundCol, startRowNum, lastRoundCol, endRowNum
				, CellRangeOptions.FixColumn | CellRangeOptions.FixRow);
			string teamRange = Utilities.CreateCellRangeString(_helper.TeamNameColumnName, startRowNum, _helper.TeamNameColumnName, endRowNum
				, CellRangeOptions.FixColumn | CellRangeOptions.FixRow);
			if (isForWinner)
			{
				// to display the winner (when all teams have taken 3 rounds)
				rule.Condition.Values.Add(new ConditionValue
				{
					// =AND($G3=1,COUNT($B$3:$E$12)=COUNT($A$3:$A$12)*3)
					// Rank == 1 && # of shootout score cells with values = (#) of teams * 3	
					UserEnteredValue = $"=AND({rankCell}=1,COUNT({scoreRange})=COUNTA({teamRange})*3)"
				});
				rule.Format.TextFormat = new TextFormat
				{
					Bold = true,
				};
			}
			else
			{
				// to display the leader (before a winner is determined)
				string teamCellColumnName = _helper.GetColumnNameByHeader(Constants.HDR_TEAM_NAME);
				string teamCellRange = Utilities.CreateCellRangeString(teamCellColumnName, startRowNum, endRowNum,
					CellRangeOptions.FixColumn | CellRangeOptions.FixRow);
				rule.Condition.Values.Add(new ConditionValue
				{
					// =AND($G3=1,COUNTBLANK($B$3:$E$12)>0,COUNTBLANK($B$3:$E$12)<COUNT($A$3:$A$12)*3)
					// Rank == 1 && # of blank shootout score cells > 0 && # of blank shootout score cells < (# of teams * 3)
					UserEnteredValue = $"=AND({rankCell}=1,COUNTBLANK({scoreRange})>0,COUNTBLANK({scoreRange})<COUNTA({teamCellRange})*3)",
				});
			}
			sheetRequests.UpdateSheetRequests.Add(request);
		}
	}
}
