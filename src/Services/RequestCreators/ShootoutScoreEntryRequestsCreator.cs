using Google.Apis.Sheets.v4.Data;
using GoogleSheetsHelper;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	/// <summary>
	/// Creates the requests to build the scoring entry section of the Shootout sheet (score entry, tiebreakers, and sorted standings)
	/// </summary>
	/// <remarks>Need one instance per division</remarks>
	public class ShootoutScoreEntryRequestsCreator : IShootoutScoreEntryRequestsCreator
	{
		private readonly ShootoutSheetConfig _shootoutSheetConfig;
		private readonly ShootoutSheetHelper _helper;
		private readonly PsoDivisionSheetHelper _divisionSheetHelper;
		private readonly string _divisionName;
		private readonly IStandingsRequestCreatorFactory _requestCreatorFactory;
		private readonly IShootoutSortedStandingsListRequestCreator _sortedStandingsCreator;
		private readonly ShootoutScoringFormulaGenerator _formulaGenerator;
		private readonly DivisionSheetConfig _divisionConfig;

		public ShootoutScoreEntryRequestsCreator(ShootoutSheetConfig shootoutSheetConfig, StandingsSheetHelper helper, PsoDivisionSheetHelper divSheetHelper
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
			// header row
			CreateScoreEntryHeaderRow(startRowIndex, ret);
			int scoreEntryStartRowIdx = startRowIndex + 1;

			// score entry rows -- one set per round of games that are followed by shootouts
			// we have to build these horizontally by pool, ie, round 1/2/3 for first pool, then same for remaining pool (if any)
			List<Tuple<int, int>> shootoutEntryStartAndEndRows;
			CreateScoreEntryRequests(info, ret, scoreEntryStartRowIdx, out shootoutEntryStartAndEndRows);

			for (int i = 0; i < _helper.DivisionSheetConfig.NumberOfShootoutRounds; i++)
			{
				int roundNum = i + 1;
				CreateScoreEntryResizeRequests(ret, roundNum);
			}

			CreateScoreDisplayRequests(ret, info, shootoutEntryStartAndEndRows);

			// tiebreaker columns
			CreateTiebreakerColumns(startRowIndex, info, ret, shootoutEntryStartAndEndRows);

			// sorted standings list
			Tuple<int, int> shootoutStartAndEnd = new(scoreEntryStartRowIdx + 1, scoreEntryStartRowIdx + _helper.DivisionSheetConfig.NumberOfTeams);
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
			for (int i = 0; i < _helper.DivisionSheetConfig.NumberOfShootoutRounds; i++)
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

		private void CreateScoreEntryRequests(PoolPlayInfo info, SheetRequests ret, int scoreEntryStartRowIdx, out List<Tuple<int, int>> shootoutEntryStartAndEndRows)
		{
			shootoutEntryStartAndEndRows = new(_helper.DivisionSheetConfig.NumberOfPools);
			int idxCounter = scoreEntryStartRowIdx;
			foreach (List<Tuple<int, int>> scoreEntryStartAndEnd in info.ScoreEntryStartAndEndRowNums.Select(p => p.Value))
			{
				for (int i = 0; i < _helper.DivisionSheetConfig.NumberOfShootoutRounds; i++)
				{
					Tuple<int, int> scoreStartAndEnd = scoreEntryStartAndEnd[i];
					Tuple<int, int> startAndEnd = CreateScoreEntryRowsForRound(idxCounter, i + 1, ret, scoreStartAndEnd);
					// since we are double looping, we need to make sure we don't add the same start/end row twice
					if (!shootoutEntryStartAndEndRows.Contains(startAndEnd))
						shootoutEntryStartAndEndRows.Add(startAndEnd);
				}
				idxCounter += _helper.DivisionSheetConfig.GamesPerRound;
			}
		}

		private Tuple<int, int> CreateScoreEntryRowsForRound(int startRowIndex, int roundNum, SheetRequests sheetRequests, Tuple<int, int> scoreEntryStartAndEnd)
		{
			// home team, home score, away score, away team
			int scoreEntryRowNum = scoreEntryStartAndEnd.Item1;
			string homeTeamNameCell = Utilities.CreateCellReference(_divisionSheetHelper.HomeTeamColumnName, scoreEntryRowNum, _divisionName);
			string homeTeamFormula = $"={homeTeamNameCell}";
			string awayTeamNameCell = Utilities.CreateCellReference(_divisionSheetHelper.AwayTeamColumnName, scoreEntryRowNum, _divisionName);
			string awayTeamFormula = $"={awayTeamNameCell}";

			Request homeTeamReq = RequestCreator.CreateRepeatedSheetFormulaRequest(_shootoutSheetConfig.SheetId, startRowIndex
				, _helper.GetColumnIndexByHeader(Constants.HDR_HOME_TEAM, roundNum), 2, homeTeamFormula);

			Request awayTeamReq = RequestCreator.CreateRepeatedSheetFormulaRequest(_shootoutSheetConfig.SheetId, startRowIndex
				, _helper.GetColumnIndexByHeader(Constants.HDR_AWAY_TEAM, roundNum), 2, awayTeamFormula);

			sheetRequests.UpdateSheetRequests.Add(homeTeamReq);
			sheetRequests.UpdateSheetRequests.Add(awayTeamReq);

			// return the start and end row of the score entry rows for this round
			return new Tuple<int, int>(startRowIndex + 1, startRowIndex + _helper.DivisionSheetConfig.GamesPerRound);
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

		private void CreateScoreDisplayRequests(SheetRequests ret, PoolPlayInfo info, List<Tuple<int, int>> shootoutEntryStartAndEndRows)
		{
			int idxCounter = shootoutEntryStartAndEndRows[0].Item1 - 1; // -1 because the tuple has row numbers
			for (int i = 0; i < _helper.DivisionSheetConfig.NumberOfPools; i++)
			{
				var poolTeams = info.Pools!.ElementAt(i);
				for (int j = 0; j < _helper.DivisionSheetConfig.NumberOfShootoutRounds; j++)
				{
					int roundNum = j + 1;
					CreateScoreDisplayColumnsForRound(poolTeams.First().TeamSheetCell, idxCounter, roundNum, ret, shootoutEntryStartAndEndRows[i]);
				}
				idxCounter += _helper.DivisionSheetConfig.TeamsPerPool;
			}
		}

		private void CreateScoreDisplayColumnsForRound(string firstTeamSheetCell, int startRowIndex, int roundNum, SheetRequests sheetRequests, Tuple<int, int> scoreEntryStartAndEnd)
		{
			int startRowNum = scoreEntryStartAndEnd.Item1;
			int endRowNum = scoreEntryStartAndEnd.Item2;
			int colIdx = _helper.GetColumnIndexByHeader($"R{roundNum}");
			// Use the 5-teams formula for divisions where not all teams play in every round
			// This happens when there are more teams per pool than games per round * 2
			string formula = _divisionConfig.TeamsPerPool > _divisionConfig.GamesPerRound * 2
				? _formulaGenerator.GetScoreDisplayFormulaWithBye(firstTeamSheetCell, startRowNum, endRowNum, roundNum)
				: _formulaGenerator.GetScoreDisplayFormula(firstTeamSheetCell, startRowNum, endRowNum, roundNum);

			Request req = RequestCreator.CreateRepeatedSheetFormulaRequest(_shootoutSheetConfig.SheetId, startRowIndex, colIdx, _divisionConfig.TeamsPerPool, formula);
			sheetRequests.UpdateSheetRequests.Add(req);
		}

		private void CreateTiebreakerColumns(int startRowIndex, PoolPlayInfo info, SheetRequests sheetRequests, List<Tuple<int, int>> shootoutEntryStartAndEndRows)
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
			int startGamesRowNum = startRowIndex + 2; // First team row number on shootout sheet
			
			// If ScoreEntryStartAndEndRowNums is populated, create separate requests per pool
			// Otherwise fall back to single request for all teams (for backward compatibility with tests)
			if (info.ScoreEntryStartAndEndRowNums.Any())
			{
				// Each pool's tiebreaker formulas need to reference that pool's division sheet rows
				int poolIdx = 0;
				foreach (var poolEntry in info.ScoreEntryStartAndEndRowNums)
				{
					// Calculate the starting row for this pool's teams on the shootout sheet
					int poolStartRowNum = startGamesRowNum + (poolIdx * _helper.DivisionSheetConfig.TeamsPerPool);
					int poolStartRowIdx = poolStartRowNum - 1; // Convert row number to index
					
					// Get this pool's division sheet row ranges for all rounds
					List<Tuple<int, int>> poolDivisionSheetRows = poolEntry.Value;
					
					foreach (string hdr in ShootoutSheetHelper.TiebreakerColumns)
					{
						IStandingsRequestCreator creator = _requestCreatorFactory.GetRequestCreator(hdr);
						ShootoutTiebreakerRequestCreatorConfig config = new()
						{
							SheetId = _shootoutSheetConfig.SheetId,
							SheetStartRowIndex = poolStartRowIdx,
							RowCount = _helper.DivisionSheetConfig.TeamsPerPool,
							StartGamesRowNum = poolStartRowNum,
							FirstTeamsSheetCell = Utilities.CreateCellReference(_helper.TeamNameColumnName, poolStartRowNum),
							ScoreEntryStartAndEndRowNums = poolDivisionSheetRows
						};
						Request request = creator.CreateRequest(config);
						sheetRequests.UpdateSheetRequests.Add(request);
					}

					// resize the columns (only need to do this once, not per pool)
					if (poolIdx == 0)
					{
						foreach (string hdr in ShootoutSheetHelper.TiebreakerColumns)
						{
							Request resizeRequest = RequestCreator.CreateCellWidthRequest(_shootoutSheetConfig.SheetId, Constants.WIDTH_WIDE_NUM_COL, _helper.GetColumnIndexByHeader(hdr));
							sheetRequests.UpdateSheetRequests.Add(resizeRequest);
						}
					}
					
					poolIdx++;
				}
			}
			else
			{
				// Fall back to original behavior: single request for all teams
				foreach (string hdr in ShootoutSheetHelper.TiebreakerColumns)
				{
					IStandingsRequestCreator creator = _requestCreatorFactory.GetRequestCreator(hdr);
					ShootoutTiebreakerRequestCreatorConfig config = new()
					{
						SheetId = _shootoutSheetConfig.SheetId,
						SheetStartRowIndex = startRowIndex + 1,
						RowCount = _helper.DivisionSheetConfig.NumberOfTeams,
						StartGamesRowNum = startGamesRowNum,
						FirstTeamsSheetCell = Utilities.CreateCellReference(_helper.TeamNameColumnName, startGamesRowNum),
						ScoreEntryStartAndEndRowNums = shootoutEntryStartAndEndRows
					};
					Request request = creator.CreateRequest(config);
					sheetRequests.UpdateSheetRequests.Add(request);

					// resize the column
					Request resizeRequest = RequestCreator.CreateCellWidthRequest(_shootoutSheetConfig.SheetId, Constants.WIDTH_WIDE_NUM_COL, _helper.GetColumnIndexByHeader(hdr));
					sheetRequests.UpdateSheetRequests.Add(resizeRequest);
				}
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
				RowCount = _helper.DivisionSheetConfig.NumberOfTeams,
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
			string lastRoundCol = _divisionConfig.NumberOfShootoutRounds == 3
				? _helper.GetColumnNameByHeader(ShootoutConstants.HDR_ROUND3)
				: _helper.GetColumnNameByHeader(ShootoutConstants.HDR_ROUND4);
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
