using Google.Apis.Sheets.v4.Data;
using GoogleSheetsHelper;
using Microsoft.Extensions.Logging;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	/// <summary>
	/// Creates the sheet for the shootout portion of the tournament: team lists,score display columns, and conditional formatting for winners.
	/// Requests for the actual score entry section, the tiebreaker columns, and sorted standings list are done in <see cref="ShootoutScoreEntryService"/>.
	/// </summary>
	public class ShootoutSheetService : IShootoutSheetService
	{
		public static string[] GetHeaderRowColumns(int numTeams) => (numTeams % 5) == 0
			? ShootoutSheetHelper.HeaderRowColumns4Rounds
			: ShootoutSheetHelper.HeaderRowColumns3Rounds;

		private readonly ISheetsClient _sheetsClient;
		private readonly ILogger<ShootoutSheetService> _logger;
		private readonly StandingsSheetHelper _helper;

		public ShootoutSheetService(ISheetsClient sheetsClient, ILogger<ShootoutSheetService> logger)
		{
			_sheetsClient = sheetsClient;
			_logger = logger;
			_helper = new StandingsSheetHelper(ShootoutSheetHelper.HeaderRowColumns4Rounds); // for the purposes in this class, it doesn't really matter
		}

		public async Task<ShootoutSheetConfig> GenerateSheet(IDictionary<string, IEnumerable<Team>> allTeams)
		{
			_logger.LogInformation("Generating Shootout sheet");

			// rename the first sheet from the default "Sheet1" -- do this first
			Sheet sheet = await RenameAndWidenSheet();
			ShootoutSheetConfig config = new()
			{
				SheetId = sheet.Properties.SheetId,
			};

			List<Request> updateSheetRequests = new List<Request>();
			List<AppendRequest> appendRequests = new List<AppendRequest>(allTeams.Sum(x => x.Value.Count() + 2));

			// create the team rows
			int rowIndex = 2; // first row that has a formula
			foreach (KeyValuePair<string, IEnumerable<Team>> division in allTeams)
			{
				List<Team> teams = division.Value.ToList();
				string[] headerRowColumns = GetHeaderRowColumns(division.Value.Count());

				AppendRequest divisionRequest = new AppendRequest(ShootoutConstants.SHOOTOUT_SHEET_NAME);

				// header row
				GoogleSheetRow headerRow = new GoogleSheetRow
				{
					new GoogleSheetCell(division.Key).SetHeaderCellFormatting()
				};
				headerRow.AddRange(Enumerable.Repeat(string.Empty, headerRowColumns.Length - 1).Select(x => new GoogleSheetCell(x).SetHeaderCellFormatting()));
				divisionRequest.Rows.Add(headerRow);

				// subheader row
				GoogleSheetRow subheaderRow = _helper.CreateHeaderRow(headerRowColumns, cell => cell.SetSubheaderCellFormatting());
				divisionRequest.Rows.Add(subheaderRow);

				int firstTeamRowNum = rowIndex + 1;
				foreach (Team team in teams)
				{
					GoogleSheetRow teamRow = CreateTeamRow(team, firstTeamRowNum, rowIndex, teams.Count);
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

		internal GoogleSheetRow CreateTeamRow(Team team, int firstTeamRowNum, int rowIndex, int teamsCount)
		{
			int rowNum = rowIndex + 1;
			team.TeamSheetCell = $"A{rowNum}";
			// shootout total: =SUM(B3:D3) -- doesn't matter if the division only has 3 rounds; the R4 column will be blank in that case
			string r1ColName = _helper.GetColumnNameByHeader(ShootoutConstants.HDR_ROUND1);
			string r4ColName = _helper.GetColumnNameByHeader(ShootoutConstants.HDR_ROUND4);
			string shootoutSumRange = Utilities.CreateCellRangeString(r1ColName, rowNum, r4ColName, rowNum, CellRangeOptions.None);
			string sumFormula = $"=SUM({shootoutSumRange})";

			// we can't do the rank formula until after the score entry section is done

			GoogleSheetRow teamRow = new GoogleSheetRow
			{
				// we are going to come back later to add formulas for the score columns
				new GoogleSheetCell(team.TeamName),
				new GoogleSheetCell(string.Empty),
				new GoogleSheetCell(string.Empty),
				new GoogleSheetCell(string.Empty),
				new GoogleSheetCell(string.Empty),
				new GoogleSheetCell() { FormulaValue = sumFormula },
			};
			return teamRow;
		}
	}
}
