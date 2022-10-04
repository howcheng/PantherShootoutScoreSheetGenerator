using Google.Apis.Sheets.v4.Data;
using GoogleSheetsHelper;
using Microsoft.Extensions.Logging;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public class ShootoutSheetService : IShootoutSheetService
	{
		private const string HDR_ROUND1 = "R1";
		private const string HDR_ROUND2 = "R2";
		private const string HDR_ROUND3 = "R3";
		public static string[] HeaderRowColumns = new string[]
		{
			Constants.HDR_TEAM_NAME,
			HDR_ROUND1,
			HDR_ROUND2,
			HDR_ROUND3,
			Constants.HDR_TOTAL_PTS,
			Constants.HDR_RANK,
		};

		private readonly ISheetsClient _sheetsClient;
		private readonly ILogger<ShootoutSheetService> _logger;
		private readonly StandingsSheetHelper _helper;
		private readonly FormulaGenerator _formulaGenerator;

		public ShootoutSheetService(ISheetsClient sheetsClient, ILogger<ShootoutSheetService> logger)
		{
			_sheetsClient = sheetsClient;
			_logger = logger;
			_helper = new StandingsSheetHelper(HeaderRowColumns);
			_formulaGenerator = new FormulaGenerator(_helper);
		}

		public async Task GenerateSheet(IDictionary<string, IEnumerable<Team>> allTeams)
		{
			_logger.LogInformation("Generating Shootout sheet");

			// rename the first sheet from the default "Sheet1"
			Sheet sheet = await _sheetsClient.GetOrAddSheet("Sheet1");
			sheet.Properties.Title = ShootoutConstants.SHOOTOUT_SHEET_NAME;
			Request titleRequest = new Request
			{
				UpdateSheetProperties = new UpdateSheetPropertiesRequest
				{
					Properties = sheet.Properties,
					Fields = "title",
				},
			};
			List<Request> updateSheetRequests = new List<Request> { titleRequest };
			List<AppendRequest> appendRequests = new List<AppendRequest>(allTeams.Sum(x => x.Value.Count() + 2));

			// create the team rows
			int rowIndex = 2; // first row that has a formula
			foreach (KeyValuePair<string, IEnumerable<Team>> division in allTeams)
			{
				AppendRequest divisionRequest = new AppendRequest(ShootoutConstants.SHOOTOUT_SHEET_NAME);

				// header row
				GoogleSheetRow headerRow = new GoogleSheetRow
				{
					new GoogleSheetCell(division.Key).SetHeaderCellFormatting()
				};
				headerRow.AddRange(Enumerable.Repeat(string.Empty, HeaderRowColumns.Count() - 1).Select(x => new GoogleSheetCell(x).SetHeaderCellFormatting()));
				divisionRequest.Rows.Add(headerRow);

				// subheader row
				GoogleSheetRow subheaderRow = new GoogleSheetRow();
				subheaderRow.AddRange(HeaderRowColumns.Select(x => new GoogleSheetCell(x).SetSubheaderCellFormatting()));
				divisionRequest.Rows.Add(subheaderRow);

				List<Team> teams = division.Value.ToList();
				int firstTeamRowNum = rowIndex + 1;
				foreach (Team team in teams)
				{
					GoogleSheetRow teamRow = CreateTeamRow(team, firstTeamRowNum, rowIndex, teams.Count);
					divisionRequest.Rows.Add(teamRow);
					rowIndex += 1;
				}

				// set up conditional formatting requests to show the leader/winner
				Request leaderFormatRequest = CreateShootoutWinnerFormatRequest(teams, sheet.Properties.SheetId, rowIndex, false);
				updateSheetRequests.Add(leaderFormatRequest);
				Request winnerFormatRequest = CreateShootoutWinnerFormatRequest(teams, sheet.Properties.SheetId, rowIndex, true);
				updateSheetRequests.Add(winnerFormatRequest);

				rowIndex += + 2; // +2 for the next set of headers
				firstTeamRowNum += teams.Count + 2;
				appendRequests.Add(divisionRequest);
			}

			await _sheetsClient.Append(appendRequests);
			await _sheetsClient.ExecuteRequests(updateSheetRequests);
		}

		internal GoogleSheetRow CreateTeamRow(Team team, int firstTeamRowNum, int rowIndex, int teamsCount)
		{
			int rowNum = rowIndex + 1;
			// shootout total: =SUM(B3:D3)
			string shootoutSumRange = Utilities.CreateCellRangeString(_helper.GetColumnNameByHeader(HDR_ROUND1), rowNum, _helper.GetColumnNameByHeader(HDR_ROUND3), rowNum, CellRangeOptions.None);
			string sumFormula = $"=SUM({shootoutSumRange})";

			// rank: =RANK(E3,E$3:E$10)
			string rankFormula = _formulaGenerator.GetTeamRankFormula(rowNum, firstTeamRowNum, firstTeamRowNum + teamsCount - 1);

			GoogleSheetRow teamRow = new GoogleSheetRow
			{
				new GoogleSheetCell(team.TeamName),
				new GoogleSheetCell(string.Empty),
				new GoogleSheetCell(string.Empty),
				new GoogleSheetCell(string.Empty),
				new GoogleSheetCell() { FormulaValue = sumFormula },
				new GoogleSheetCell() { FormulaValue = rankFormula },
			};
			return teamRow;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="teams"></param>
		/// <param name="sheetId"></param>
		/// <param name="startRowIndex"></param>
		/// <param name="isForWinner"><c>false</c> for the leader (i.e., formatting before a winner is determined)</param>
		/// <returns></returns>
		internal Request CreateShootoutWinnerFormatRequest(List<Team> teams, int? sheetId, int startRowIndex, bool isForWinner)
		{
			int startRowNum = startRowIndex + 1;
			int endRowNum = startRowIndex + teams.Count;
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
								SheetId = sheetId,
								StartRowIndex = startRowIndex,
								StartColumnIndex = _helper.GetColumnIndexByHeader(Constants.HDR_TEAM_NAME),
								EndRowIndex = startRowIndex + teams.Count,
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
			string scoreRange = Utilities.CreateCellRangeString(_helper.GetColumnNameByHeader(HDR_ROUND1), startRowNum, _helper.GetColumnNameByHeader(HDR_ROUND3), endRowNum,
				CellRangeOptions.FixColumn | CellRangeOptions.FixRow);
			if (isForWinner)
			{
				rule.Condition.Values.Add(new ConditionValue
				{
					// =AND($F3=1,COUNTBLANK($B$3:$D$14)=0)
					// Rank == 1 && # of blank shootout score cells == 0
					UserEnteredValue = $"=AND({rankCell}=1,COUNTBLANK({scoreRange})=0)"
				});
				rule.Format.TextFormat = new TextFormat
				{
					Bold = true,
				};
			}
			else
			{
				string teamCellColumnName = _helper.GetColumnNameByHeader(Constants.HDR_TEAM_NAME);
				string teamCellRange = Utilities.CreateCellRangeString(teamCellColumnName, startRowNum, endRowNum,
					CellRangeOptions.FixColumn | CellRangeOptions.FixRow);
				rule.Condition.Values.Add(new ConditionValue
				{
					// =AND($F3=1,COUNTBLANK($B$3:$D$14)>0,COUNTBLANK($B$3:$D$14)<COUNTA($A$3:$A$14)*3)
					// Rank == 1 && # of blank shootout score cells > 0 && # of blank shootout score cells < (# of teams * 3)
					UserEnteredValue = $"=AND({rankCell}=1,COUNTBLANK({scoreRange})>0,COUNTBLANK({scoreRange})<COUNTA({teamCellRange})*3)",
				});
			}
			return request;
		}
	}
}
