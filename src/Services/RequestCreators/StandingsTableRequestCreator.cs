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
		private readonly PsoDivisionSheetHelper _helper;
		private readonly IStandingsRequestCreatorFactory _requestCreatorFactory;

		public StandingsTableRequestCreator(PsoDivisionSheetHelper helper, IStandingsRequestCreatorFactory factory)
		{
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
		public PoolPlayInfo CreateStandingsRequests(DivisionSheetConfig config, PoolPlayInfo info, IEnumerable<Team> poolTeams, int startRowIndex)
		{
			List<Request> requests = new List<Request>(_helper.StandingsTableColumns.Count + poolTeams.Count() * 2);
			int teamCount = poolTeams.Count();
			string firstTeamSheetCell = $"{ShootoutConstants.SHOOTOUT_SHEET_NAME}!{poolTeams.First().TeamSheetCell}";

			// team name
			Request teamNamesRequest = RequestCreator.CreateRepeatedSheetFormulaRequest(config.SheetId, startRowIndex, _helper.GetColumnIndexByHeader(Constants.HDR_TEAM_NAME), teamCount,
				$"={firstTeamSheetCell}");
			requests.Add(teamNamesRequest);

			int startGamesRowNum = startRowIndex + 1; // first row in first round is 3
			int endGamesRowNum = GetEndGamesInPoolRowNumber(config, startGamesRowNum);

			foreach (string hdr in _helper.StandingsTableColumns)
			{
				IStandingsRequestCreator requestCreator = _requestCreatorFactory.GetRequestCreator(hdr);
				if (requestCreator == null) // not all columns have formulas (yellow/red cards, tiebreaker)
					continue;

				PsoStandingsRequestCreatorConfig creatorConfig = new PsoStandingsRequestCreatorConfig
				{
					StartGamesRowNum = startGamesRowNum,
					EndGamesRowNum = endGamesRowNum,
					FirstTeamsSheetCell = firstTeamSheetCell,
					GamesPerRound = config.GamesPerRound,
					NumberOfRounds = config.NumberOfRounds,
					RowCount = poolTeams.Count(),
					SheetId = config.SheetId,
				};
				Request request = requestCreator.CreateRequest(creatorConfig);
				info.UpdateSheetRequests.Add(request);
			}

			// head-to-head tiebreakers

			Team firstTeam = poolTeams.First();
			System.Text.RegularExpressions.Regex reTeamCell = new System.Text.RegularExpressions.Regex(@"(A)(\d+)");
			int startColumnIndex = _helper.GetColumnIndexByHeader(_helper.StandingsTableColumns.Last()) + 1;
			for (int i = 0; i < teamCount; i++)
			{
				IStandingsRequestCreator requestCreator = _requestCreatorFactory.GetRequestCreator(nameof(HeadToHeadComparisonRequestCreator));
				Team team = poolTeams.ElementAt(i);
				// we have to fix the opponent team cell
				System.Text.RegularExpressions.GroupCollection groups = reTeamCell.Match(team.TeamSheetCell).Groups;
				string opponentCell = $"{groups[1].Value}${groups[2].Value}";

				HeadToHeadComparisonRequestCreatorConfig headToHeadConfig = new HeadToHeadComparisonRequestCreatorConfig
				{
					SheetId = config.SheetId,
					ColumnIndex = startColumnIndex + i,
					SheetStartRowIndex = startRowIndex,
					RowCount = teamCount,
					StartGamesRowNum = startGamesRowNum,
					EndGamesRowNum = endGamesRowNum,
					FirstTeamsSheetCell = firstTeam.TeamSheetCell,
					OpponentTeamSheetCell = opponentCell,
				};

				Request headToHeadRequest = requestCreator.CreateRequest(headToHeadConfig);
				requests.Add(headToHeadRequest);
			}

			// resize the head-to-head columns
			Request resizeRequest = new Request
			{
				UpdateDimensionProperties = new UpdateDimensionPropertiesRequest
				{
					Range = new DimensionRange
					{
						Dimension = "COLUMNS",
						SheetId = config.SheetId,
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
			requests.Add(resizeRequest);
			info.UpdateSheetRequests.AddRange(requests);
			return info;
		}

		/// <summary>
		/// For formulas that require a cell range to cover all games, gets the row number of the last game
		/// </summary>
		/// <param name="startGamesRowNum"></param>
		/// <returns></returns>
		private int GetEndGamesInPoolRowNumber(DivisionSheetConfig config, int startGamesRowNum)
		{
			int offset = (config.NumberOfRounds * config.GamesPerRound) + (config.NumberOfRounds - 1 * Constants.ROUND_OFFSET_STANDINGS_TABLE);
			return startGamesRowNum + offset - 1; // if we have 10 rows total (including blank/header rows), A3:A12 includes 10 rows so -1 to make it work
		}
	}
}
