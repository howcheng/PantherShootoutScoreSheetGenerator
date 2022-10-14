using Google.Apis.Sheets.v4.Data;
using GoogleSheetsHelper;
using StandingsGoogleSheetsHelper;
using StandingsConstants = StandingsGoogleSheetsHelper.Constants;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public abstract class DivisionSheetCreator : IDivisionSheetCreator
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

		protected DivisionSheetCreator(DivisionSheetConfig config, ISheetsClient sheetsClient, FormulaGenerator fg
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
		public static List<string> HeaderRowColumns = new List<string>
		{
			StandingsConstants.HDR_HOME_TEAM,
			StandingsConstants.HDR_HOME_GOALS,
			StandingsConstants.HDR_AWAY_GOALS,
			StandingsConstants.HDR_AWAY_TEAM,
		};
		// PSO winner column is moved off to the side so that it's not visible when the sheet is embedded
		public static List<string> WinnerAndPointsColumns = new List<string>
		{
			StandingsConstants.HDR_WINNING_TEAM,
			StandingsConstants.HDR_HOME_PTS,
			StandingsConstants.HDR_AWAY_PTS,
		};
		public static List<string> StandingsHeaderRow = new List<string>
		{
			StandingsConstants.HDR_TEAM_NAME,
			StandingsConstants.HDR_GAMES_PLAYED,
			StandingsConstants.HDR_NUM_WINS,
			StandingsConstants.HDR_NUM_LOSSES,
			StandingsConstants.HDR_NUM_DRAWS,
			StandingsConstants.HDR_YELLOW_CARDS,
			StandingsConstants.HDR_RED_CARDS,
			StandingsConstants.HDR_GAME_PTS,
			StandingsConstants.HDR_RANK,
			StandingsConstants.HDR_CALC_RANK,
			StandingsConstants.HDR_TIEBREAKER,
			StandingsConstants.HDR_GOALS_FOR,
			StandingsConstants.HDR_GOALS_AGAINST,
			StandingsConstants.HDR_GOAL_DIFF,
		};
		protected static string END_STANDINGS_COL = StandingsConstants.HDR_GOAL_DIFF;
		#endregion

		public async Task CreateSheet()
		{
			if (!SheetId.HasValue)
			{
				Sheet sheet = await _sheetsClient.GetOrAddSheet(DivisionName);
				SheetId = sheet.Properties.SheetId;
			}

			// pool play updates
			PoolPlayInfo poolPlay = new PoolPlayInfo(_divisionTeams.GroupBy(x => x.PoolName).OrderBy(x => x.Key));
			poolPlay = await _poolPlayRequestCreator.CreatePoolPlayRequests(poolPlay);
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
