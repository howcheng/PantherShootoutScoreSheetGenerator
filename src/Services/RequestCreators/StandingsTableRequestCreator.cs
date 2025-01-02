using Google.Apis.Sheets.v4.Data;
using GoogleSheetsHelper;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	/// <summary>
	/// Creates the requests for the standings table; not to be confused with classes that implement  <see cref="IStandingsRequestCreator"/>
	/// </summary>
	public class StandingsTableRequestCreator : IStandingsTableRequestCreator
	{
		private static readonly System.Text.RegularExpressions.Regex s_reTeamCell = new System.Text.RegularExpressions.Regex(@"(A)(\d+)");

		protected readonly DivisionSheetConfig _config;
		protected readonly PsoDivisionSheetHelper _helper;
		protected readonly IStandingsRequestCreatorFactory _requestCreatorFactory;

		public StandingsTableRequestCreator(DivisionSheetConfig config, PsoDivisionSheetHelper helper, IStandingsRequestCreatorFactory factory)
		{
			_config = config;
			_helper = helper;
			_requestCreatorFactory = factory;
		}

		/// <summary>
		/// Creates the <see cref="Request"/> and <see cref="UpdateRequest"/> objects to build the standings section of the sheet
		/// </summary>
		/// <param name="info"></param>
		/// <param name="poolTeams"></param>
		/// <param name="startRowIndex"></param>
		/// <returns></returns>
		public virtual PoolPlayInfo CreateStandingsRequests(PoolPlayInfo info, IEnumerable<Team> poolTeams, int startRowIndex)
		{
			List<Request> requests = new List<Request>(_helper.StandingsTableColumns.Count + poolTeams.Count());

			// team name
			Team firstTeam = poolTeams.First();
			string firstTeamSheetCell = $"{ShootoutConstants.SHOOTOUT_SHEET_NAME}!{firstTeam.TeamSheetCell}";
			Request teamNamesRequest = RequestCreator.CreateRepeatedSheetFormulaRequest(_config.SheetId, startRowIndex, _helper.GetColumnIndexByHeader(Constants.HDR_TEAM_NAME), _config.TeamsPerPool,
				$"={firstTeamSheetCell}");
			requests.Add(teamNamesRequest);

			int startGamesRowNum = startRowIndex + 1; // first row in first round is 3
			int endGamesRowNum = GetEndGamesInPoolRowNumber(_config, startGamesRowNum);

			// standings table
			foreach (string hdr in _helper.StandingsTableColumns)
			{
				Request? request = CreateRequestForStandingsTableColumn(hdr, startGamesRowNum, startRowIndex, endGamesRowNum, firstTeamSheetCell);
				if (request != null)
					requests.Add(request);
			}

			// head-to-head tiebreakers
			int startColumnIndex = _helper.GetColumnIndexByHeader(_helper.StandingsTableColumns.Last()) + 1;
			for (int i = 0; i < _config.TeamsPerPool; i++)
			{
				Request headToHeadRequest = CreateH2HTiebreakerColumnRequest(poolTeams, startRowIndex, firstTeamSheetCell, startGamesRowNum, endGamesRowNum, startColumnIndex, i);
				requests.Add(headToHeadRequest);
			}

			// resize the head-to-head columns
			Request resizeRequest = CreateH2HResizeRequest(poolTeams, startColumnIndex);
			requests.Add(resizeRequest);
			info.UpdateSheetRequests.AddRange(requests);
			return info;
		}

		protected virtual Request? CreateRequestForStandingsTableColumn(string hdr, int startGamesRowNum, int startRowIndex, int endGamesRowNum, string firstTeamSheetCell)
		{
			IStandingsRequestCreator requestCreator = _requestCreatorFactory.GetRequestCreator(hdr);
			if (requestCreator == null) // not all columns have formulas (yellow/red cards)
				return null;

			PsoStandingsRequestCreatorConfig creatorConfig = new PsoStandingsRequestCreatorConfig
			{
				SheetId = _config.SheetId,
				StartGamesRowNum = startGamesRowNum,
				SheetStartRowIndex = startRowIndex,
				EndGamesRowNum = endGamesRowNum,
				FirstTeamsSheetCell = firstTeamSheetCell,
				GamesPerRound = _config.GamesPerRound,
				NumberOfRounds = _config.NumberOfRounds,
				RowCount = _config.TeamsPerPool,
			};
			Request request = requestCreator.CreateRequest(creatorConfig);
			return request;
		}

		private Request CreateH2HTiebreakerColumnRequest(IEnumerable<Team> poolTeams, int startRowIndex, string firstTeamSheetCell, int startGamesRowNum, int endGamesRowNum, int startColumnIndex, int i)
		{
			IStandingsRequestCreator requestCreator = _requestCreatorFactory.GetRequestCreator(nameof(HeadToHeadComparisonRequestCreator));
			Team team = poolTeams.ElementAt(i);
			// we have to fix the opponent team cell
			System.Text.RegularExpressions.GroupCollection groups = s_reTeamCell.Match(team.TeamSheetCell).Groups;
			string opponentCell = $"{ShootoutConstants.SHOOTOUT_SHEET_NAME}!{groups[1].Value}${groups[2].Value}";

			HeadToHeadComparisonRequestCreatorConfig headToHeadConfig = new HeadToHeadComparisonRequestCreatorConfig
			{
				SheetId = _config.SheetId,
				ColumnIndex = startColumnIndex + i,
				SheetStartRowIndex = startRowIndex,
				RowCount = _config.TeamsPerPool,
				StartGamesRowNum = startGamesRowNum,
				EndGamesRowNum = endGamesRowNum,
				FirstTeamsSheetCell = firstTeamSheetCell,
				OpponentTeamSheetCell = opponentCell,
			};

			Request headToHeadRequest = requestCreator.CreateRequest(headToHeadConfig);
			return headToHeadRequest;
		}

		private Request CreateH2HResizeRequest(IEnumerable<Team> poolTeams, int startColumnIndex) => new Request
		{
			UpdateDimensionProperties = new UpdateDimensionPropertiesRequest
			{
				Range = new DimensionRange
				{
					Dimension = "COLUMNS",
					SheetId = _config.SheetId,
					StartIndex = startColumnIndex,
					EndIndex = startColumnIndex + poolTeams.Count(),
				},
				Properties = new DimensionProperties
				{
					PixelSize = Constants.WIDTH_WIDE_NUM_COL,
				},
				Fields = "*"
			},
		};

		/// <summary>
		/// For formulas that require a cell range to cover all games, gets the row number of the last game
		/// </summary>
		/// <param name="startGamesRowNum"></param>
		/// <returns></returns>
		private static int GetEndGamesInPoolRowNumber(DivisionSheetConfig config, int startGamesRowNum)
		{
			int offset = (config.NumberOfRounds * config.GamesPerRound) + ((config.NumberOfRounds - 1) * Constants.ROUND_OFFSET_STANDINGS_TABLE);
			return startGamesRowNum + offset - 1; // if we have 10 rows total (including blank/header rows), A3:A12 includes 10 rows so -1 to make it work
		}
	}
}
