using Google.Apis.Sheets.v4.Data;
using GoogleSheetsHelper;
using Microsoft.Extensions.Logging;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	/// <summary>
	/// Creates the sheet for the shootout portion of the tournament: team lists,score display columns, and conditional formatting for winners.
	/// Requests for the actual score entry section, the tiebreakers, and sorted standings list are done in <see cref="ShootoutScoreEntryService"/>.
	/// </summary>
	public class ShootoutSheetService : IShootoutSheetService
	{
		private readonly ISheetsClient _sheetsClient;
		private readonly ILogger<ShootoutSheetService> _logger;
		private readonly StandingsSheetHelper _helper;

		public ShootoutSheetService(ISheetsClient sheetsClient, ILogger<ShootoutSheetService> logger)
		{
			_sheetsClient = sheetsClient;
			_logger = logger;
			_helper = new StandingsSheetHelper(ShootoutSheetHelper.HeaderRowColumns4Rounds); // for the purposes in this class, it doesn't really matter
		}

		public async Task<ShootoutSheetConfig> GenerateSheet(IDictionary<string, IEnumerable<Team>> allTeams, IDictionary<string, DivisionSheetConfig> divisionConfigs)
		{
			_logger.LogInformation("Generating Shootout sheet");

			// rename the first sheet from the default "Sheet1" -- do this first
			Sheet sheet = await RenameAndWidenSheet();
			ShootoutSheetConfig config = new()
			{
				SheetId = sheet.Properties.SheetId,
				DivisionConfigs = divisionConfigs.ToDictionary(kvp => kvp.Key, kvp => kvp.Value)
			};

			List<Request> updateSheetRequests = new List<Request>();
			List<AppendRequest> appendRequests = new List<AppendRequest>(allTeams.Sum(x => x.Value.Count() + 2));

			// create the team rows
			int rowIndex = 2; // first row that has a formula
			foreach (KeyValuePair<string, IEnumerable<Team>> division in allTeams)
			{
				List<Team> teams = division.Value.ToList();
				DivisionSheetConfig divisionConfig = divisionConfigs[division.Key];
				ShootoutSheetHelper divisionHelper = new ShootoutSheetHelper(divisionConfig);

				AppendRequest divisionRequest = new AppendRequest(ShootoutConstants.SHOOTOUT_SHEET_NAME);

				// header row
				GoogleSheetRow headerRow = new GoogleSheetRow
				{
					new GoogleSheetCell(division.Key).SetHeaderCellFormatting()
				};
				headerRow.AddRange(Enumerable.Repeat(string.Empty, divisionHelper.HeaderRowColumns.Count - 1).Select(x => new GoogleSheetCell(x).SetHeaderCellFormatting()));
				divisionRequest.Rows.Add(headerRow);

				// subheader row
				GoogleSheetRow subheaderRow = divisionHelper.CreateHeaderRow(divisionHelper.HeaderRowColumns, cell => cell.SetSubheaderCellFormatting());
				divisionRequest.Rows.Add(subheaderRow);

				int firstTeamRowNum = rowIndex + 1;
				foreach (Team team in teams)
				{
					GoogleSheetRow teamRow = CreateTeamRow(team, firstTeamRowNum, rowIndex, teams.Count, divisionHelper);
					divisionRequest.Rows.Add(teamRow);
					rowIndex += 1;
				}
				config.FirstTeamSheetCells.Add(division.Key, teams.First().TeamSheetCell);

				Tuple<int, int> startAndEnd = new(firstTeamRowNum, rowIndex);
				config.ShootoutStartAndEndRows.Add(division.Key, startAndEnd);

				rowIndex += +2; // +2 for the next set of headers
				firstTeamRowNum += teams.Count + 2;
				appendRequests.Add(divisionRequest);
			}

			// resize the columns (except team name, which we will do later)
			IEnumerable<Request> resizeRequests = _helper.CreateCellWidthRequests(sheet.Properties.SheetId, 0);
			updateSheetRequests.AddRange(resizeRequests.Skip(1));

			await _sheetsClient.Append(appendRequests);
			await _sheetsClient.ExecuteRequests(updateSheetRequests);

			// resize the team name column
			config.TeamNameCellWidth = await _sheetsClient.AutoResizeColumn(ShootoutConstants.SHOOTOUT_SHEET_NAME, 0);
			return config;
		}

		private async Task<Sheet> RenameAndWidenSheet()
		{
			Sheet sheet = await _sheetsClient.GetOrAddSheet("Sheet1");
			sheet.Properties.Title = ShootoutConstants.SHOOTOUT_SHEET_NAME;
			sheet.Properties.GridProperties.ColumnCount = 40;
			Request titleRequest = new Request
			{
				UpdateSheetProperties = new UpdateSheetPropertiesRequest
				{
					Properties = sheet.Properties,
					Fields = $"{nameof(SheetProperties.Title).ToLower()},{nameof(GridProperties).ToCamelCase()}({nameof(GridProperties.ColumnCount).ToCamelCase()})",
				},
			};

			await _sheetsClient.ExecuteRequests(new[] { titleRequest });
			return sheet;
		}

		internal GoogleSheetRow CreateTeamRow(Team team, int firstTeamRowNum, int rowIndex, int teamsCount, ShootoutSheetHelper helper)
		{
			int rowNum = rowIndex + 1;
			team.TeamSheetCell = $"A{rowNum}";
			// shootout total: =SUM(B3:D3) or =SUM(B3:E3) depending on number of rounds
			string r1ColName = helper.GetColumnNameByHeader(ShootoutConstants.HDR_ROUND1);
			string lastRoundColName = helper.DivisionSheetConfig.NumberOfShootoutRounds == 3
				? helper.GetColumnNameByHeader(ShootoutConstants.HDR_ROUND3)
				: helper.GetColumnNameByHeader(ShootoutConstants.HDR_ROUND4);
			string shootoutSumRange = Utilities.CreateCellRangeString(r1ColName, rowNum, lastRoundColName, rowNum, CellRangeOptions.None);
			string sumFormula = $"=SUM({shootoutSumRange})";

			// we can't do the rank formula until after the score entry section is done

			GoogleSheetRow teamRow = new GoogleSheetRow
			{
				// we are going to come back later to add formulas for the score columns
				new GoogleSheetCell(team.TeamName),
				new GoogleSheetCell(string.Empty),
				new GoogleSheetCell(string.Empty),
				new GoogleSheetCell(string.Empty),
			};
			
			// Add 4th round column if needed
			if (helper.DivisionSheetConfig.NumberOfShootoutRounds == 4)
			{
				teamRow.Add(new GoogleSheetCell(string.Empty));
			}
			
			teamRow.Add(new GoogleSheetCell() { FormulaValue = sumFormula });
			return teamRow;
		}

		/// <summary>
		/// Hides the helper columns (score entry, tiebreakers, sorted standings list) on the Shootout sheet
		/// </summary>
		/// <param name="config">The shootout sheet configuration</param>
		public async Task HideHelperColumns(ShootoutSheetConfig config)
		{
			_logger.LogInformation("Hiding helper columns on Shootout sheet");

			// Get any division config to determine the structure (they all have the same column layout)
			DivisionSheetConfig anyDivisionConfig = config.DivisionConfigs.Values.First();
			ShootoutSheetHelper helper = new ShootoutSheetHelper(anyDivisionConfig);
			IColumnVisibilityHelper columnVisibilityHelper = new ColumnVisibilityHelper();

			// Calculate the range to hide: everything after the RANK column through the end of the sorted standings list
			int startHideColumn = helper.LastVisibleColumnIndex + 1;
			// The sorted standings list width equals the number of teams in the division
			int endHideColumn = helper.SortedStandingsListColumnIndex + anyDivisionConfig.NumberOfTeams;

			IList<Request> hideColumnsRequests = columnVisibilityHelper.CreateHideColumnsRequest(config.SheetId, startHideColumn, endHideColumn);
			await _sheetsClient.ExecuteRequests(hideColumnsRequests);
		}
	}
}
