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

		#region Constants for sheet
		public static List<string> GameScoreColumns = new List<string>
		{
			Constants.HDR_HOME_TEAM,
			Constants.HDR_HOME_GOALS,
			Constants.HDR_AWAY_GOALS,
			Constants.HDR_AWAY_TEAM,
			Constants.HDR_FORFEIT,
		};
		// PSO winner column is moved off to the side so that it's not visible when the sheet is embedded
		public static List<string> WinnerAndPointsColumns = new List<string>
		{
			Constants.HDR_WINNING_TEAM,
			Constants.HDR_HOME_PTS,
			Constants.HDR_AWAY_PTS,
		};
		public static List<string> StandingsHeaderRow = new List<string>
		{
			Constants.HDR_TEAM_NAME,
			Constants.HDR_GAMES_PLAYED,
			Constants.HDR_NUM_WINS,
			Constants.HDR_NUM_LOSSES,
			Constants.HDR_NUM_DRAWS,
			Constants.HDR_YELLOW_CARDS,
			Constants.HDR_RED_CARDS,
			Constants.HDR_GAME_PTS,
			Constants.HDR_RANK,
			Constants.HDR_CALC_RANK,
			Constants.HDR_TIEBREAKER,
			Constants.HDR_GOALS_FOR,
			Constants.HDR_GOALS_AGAINST,
			Constants.HDR_GOAL_DIFF,
		};
		protected static string END_STANDINGS_COL = Constants.HDR_GOAL_DIFF;
		#endregion

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
