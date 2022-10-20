using Google.Apis.Sheets.v4.Data;
using GoogleSheetsHelper;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public class DivisionSheetGenerator : IDivisionSheetGenerator
	{
		private DivisionSheetConfig _config;
		private List<Team> _divisionTeams;

		protected int? SheetId { get; private set; }
		protected string DivisionName => _divisionTeams.First().DivisionName;

		protected readonly ISheetsClient _sheetsClient;
		private readonly IPoolPlayRequestCreator _poolPlayRequestCreator;
		private readonly IChampionshipRequestCreator _championshipRequestCreator;
		private readonly IWinnerFormattingRequestsCreator _winnerFormattingRequestCreator;

		public PsoDivisionSheetHelper Helper { get; private set; }

		public DivisionSheetGenerator(DivisionSheetConfig config, ISheetsClient sheetsClient, FormulaGenerator fg
			, IPoolPlayRequestCreator poolPlayCreator, IChampionshipRequestCreator championshipCreator, IWinnerFormattingRequestsCreator winnerFormatCreator
			, List<Team> divisionTeams)
		{
			_config = config;
			_sheetsClient = sheetsClient;
			SheetId = _config.SheetId;

			Helper = (PsoDivisionSheetHelper)fg.SheetHelper;
			_poolPlayRequestCreator = poolPlayCreator;
			_championshipRequestCreator = championshipCreator;
			_winnerFormattingRequestCreator = winnerFormatCreator;

			_divisionTeams = divisionTeams;
		}

		protected static string END_STANDINGS_COL = Constants.HDR_GOAL_DIFF;

		public async Task CreateSheet()
		{
			if (!SheetId.HasValue)
			{
				Sheet sheet = await _sheetsClient.GetOrAddSheet(DivisionName);
				SheetId = sheet.Properties.SheetId;
			}

			// pool play updates
			PoolPlayInfo poolPlay = new PoolPlayInfo(_divisionTeams);
			poolPlay = _poolPlayRequestCreator.CreatePoolPlayRequests(poolPlay);
			await _sheetsClient.Update(poolPlay.UpdateValuesRequests);
			await _sheetsClient.ExecuteRequests(poolPlay.UpdateSheetRequests);

			// resize the columns
			List<Request> resizeRequests = new List<Request>();
			resizeRequests.AddRange(Helper.CreateCellWidthRequests(SheetId, _config.TeamNameCellWidth));
			await _sheetsClient.ExecuteRequests(resizeRequests);

			// championship updates
			ChampionshipInfo championship = _championshipRequestCreator.CreateChampionshipRequests(poolPlay);
			if (championship.UpdateValuesRequests.Count > 0)
				await _sheetsClient.Update(championship.UpdateValuesRequests);

			if (championship.UpdateSheetRequests.Count > 0)
				await _sheetsClient.ExecuteRequests(championship.UpdateSheetRequests);

			// winner conditional formatting
			SheetRequests winnerRequests = _winnerFormattingRequestCreator.CreateWinnerFormattingRequests(_config, championship);
			await _sheetsClient.Update(winnerRequests.UpdateValuesRequests);
			await _sheetsClient.ExecuteRequests(winnerRequests.UpdateSheetRequests);
		}
	}
}
