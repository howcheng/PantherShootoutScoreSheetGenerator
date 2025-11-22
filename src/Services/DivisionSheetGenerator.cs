using Google.Apis.Sheets.v4.Data;
using GoogleSheetsHelper;
using Microsoft.Extensions.Logging;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public class DivisionSheetGenerator : IDivisionSheetGenerator
	{
		private readonly DivisionSheetConfig _config;
		private readonly IEnumerable<Team> _divisionTeams;

		protected int? SheetId { get; private set; }
		protected string DivisionName => _divisionTeams.First().DivisionName;

		protected readonly ISheetsClient _sheetsClient;
		private readonly IPoolPlayRequestCreator _poolPlayRequestCreator;
		private readonly IChampionshipRequestCreator _championshipRequestCreator;
		private readonly IWinnerFormattingRequestsCreator _winnerFormattingRequestCreator;
		private readonly IColumnVisibilityHelper _columnVisibilityHelper;
		private readonly ILogger<DivisionSheetGenerator> _log;

		public PsoDivisionSheetHelper Helper { get; private set; }

		public DivisionSheetGenerator(DivisionSheetConfig config, ISheetsClient sheetsClient, FormulaGenerator fg
			, IPoolPlayRequestCreator poolPlayCreator, IChampionshipRequestCreator championshipCreator, IWinnerFormattingRequestsCreator winnerFormatCreator
			, IColumnVisibilityHelper columnVisibilityHelper
			, ILogger<DivisionSheetGenerator> log
			, IEnumerable<Team> divisionTeams)
		{
			_config = config;
			_sheetsClient = sheetsClient;
			SheetId = _config.SheetId;

			Helper = (PsoDivisionSheetHelper)fg.SheetHelper;
			_poolPlayRequestCreator = poolPlayCreator;
			_championshipRequestCreator = championshipCreator;
			_winnerFormattingRequestCreator = winnerFormatCreator;
			_columnVisibilityHelper = columnVisibilityHelper;
			_log = log;

			_divisionTeams = divisionTeams;
		}

		protected const string END_STANDINGS_COL = Constants.HDR_GOAL_DIFF;

		public async Task<PoolPlayInfo> CreateSheet(ShootoutSheetConfig shootoutSheetConfig)
		{
			_log.LogInformation($"Beginning sheet for {DivisionName}...");
			if (!SheetId.HasValue)
			{
				Sheet sheet = await _sheetsClient.GetOrAddSheet(DivisionName);
				_config.SheetId = SheetId = sheet.Properties.SheetId;
			}

			// first, expand the sheet because we be wiiiiiide (only 26 columns are available otherwise)
			Request resizeSheetRequest = new Request
			{
				UpdateSheetProperties = new UpdateSheetPropertiesRequest
				{
					Properties = new SheetProperties
					{
						GridProperties = new GridProperties
						{
							ColumnCount = Helper.SortedStandingsListColumnIndex + 30, // this is an arbitrary number because the exact number could vary depending on how many teams are in the pool
						},
						SheetId = SheetId,
					},
					Fields = $"{nameof(GridProperties).ToCamelCase()}({nameof(GridProperties.ColumnCount).ToCamelCase()})",
				}
			};
			await _sheetsClient.ExecuteRequests(new[] { resizeSheetRequest });

			// pool play updates
			PoolPlayInfo poolPlay = new PoolPlayInfo(_divisionTeams);
			poolPlay = _poolPlayRequestCreator.CreatePoolPlayRequests(poolPlay);
			await _sheetsClient.Update(poolPlay.UpdateValuesRequests);
			await _sheetsClient.ExecuteRequests(poolPlay.UpdateSheetRequests);

			// resize the columns
			List<Request> resizeRequests = new List<Request>();
			resizeRequests.AddRange(Helper.CreateCellWidthRequests(SheetId, _config.TeamNameCellWidth));
			resizeRequests.AddRange(Helper.CreateCellWidthRequests(SheetId, PsoDivisionSheetHelper.MainTiebreakerColumns, _config.TeamNameCellWidth));
			await _sheetsClient.ExecuteRequests(resizeRequests);

			// championship updates
			ChampionshipInfo championship = _championshipRequestCreator.CreateChampionshipRequests(poolPlay);
			if (championship.UpdateValuesRequests.Count > 0)
				await _sheetsClient.Update(championship.UpdateValuesRequests);

			if (championship.UpdateSheetRequests.Count > 0)
				await _sheetsClient.ExecuteRequests(championship.UpdateSheetRequests);

			// winner conditional formatting
			SheetRequests winnerRequests = _winnerFormattingRequestCreator.CreateWinnerFormattingRequests(championship);
			await _sheetsClient.Update(winnerRequests.UpdateValuesRequests);
			await _sheetsClient.ExecuteRequests(winnerRequests.UpdateSheetRequests);

			// Hide helper columns (H2H tiebreakers, main tiebreakers, winner/points, sorted standings list)
			int startHideColumn = Helper.LastVisibleColumnIndex + 1;
			int endHideColumn = Helper.SortedStandingsListColumnIndex + _divisionTeams.Count(); // +teams count for sorted list width
			IList<Request> hideColumnsRequests = _columnVisibilityHelper.CreateHideColumnsRequest(SheetId, startHideColumn, endHideColumn);
			await _sheetsClient.ExecuteRequests(hideColumnsRequests);

			// Return ChampionshipInfo (which extends PoolPlayInfo) so shootout score entry can access championship game rows
			return championship;
		}
	}
}
