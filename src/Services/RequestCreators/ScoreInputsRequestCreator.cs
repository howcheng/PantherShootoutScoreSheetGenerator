using Google.Apis.Sheets.v4.Data;
using GoogleSheetsHelper;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public class ScoreInputsRequestCreator : IScoreInputsRequestCreator
	{
		private readonly string _divisionName;
		private readonly PsoDivisionSheetHelper _helper;
		private readonly IStandingsRequestCreatorFactory _requestCreatorFactory;

		public ScoreInputsRequestCreator(string divisionName, PsoDivisionSheetHelper helper, IStandingsRequestCreatorFactory factory)
		{
			_divisionName = divisionName;
			_helper = helper;
			_requestCreatorFactory = factory;
		}

		public virtual PoolPlayInfo CreateScoringRequests(DivisionSheetConfig config, PoolPlayInfo info, IEnumerable<Team> poolTeams, int roundNum, ref int startRowIndex)
		{
			List<Request> requests = new List<Request>();

			// round label
			UpdateRequest roundHeaderRequest = Utilities.CreateHeaderLabelRowRequest(_divisionName, startRowIndex++, _helper.GetColumnIndexByHeader(Constants.HDR_AWAY_TEAM), $"ROUND {roundNum}", 0, cell => cell.SetSubheaderCellFormatting());
			info.UpdateValuesRequests.Add(roundHeaderRequest);

			// scoring rows
			int startRowNum = startRowIndex + 1;
			string firstTeamSheetCell = poolTeams.First().TeamSheetCell;
			string lastTeamSheetCell = poolTeams.Last().TeamSheetCell;
			requests.Add(RequestCreator.CreateDataValidationRequest(ShootoutConstants.SHOOTOUT_SHEET_NAME, firstTeamSheetCell, lastTeamSheetCell, config.SheetId, startRowIndex + 1, _helper.GetColumnIndexByHeader(Constants.HDR_HOME_TEAM), config.GamesPerRound));
			requests.Add(RequestCreator.CreateDataValidationRequest(ShootoutConstants.SHOOTOUT_SHEET_NAME, firstTeamSheetCell, lastTeamSheetCell, config.SheetId, startRowIndex + 1, _helper.GetColumnIndexByHeader(Constants.HDR_AWAY_TEAM), config.GamesPerRound));

			UpdateRequest winnerAndPtsHeadersRequest = new UpdateRequest(_divisionName)
			{
				RowStart = startRowIndex,
			};
			winnerAndPtsHeadersRequest.Rows.Add(_helper.CreateHeaderRow(_helper.WinnerAndPointsColumns));

			foreach (string hdr in _helper.WinnerAndPointsColumns)
			{
				IStandingsRequestCreator requestCreator = _requestCreatorFactory.GetRequestCreator(hdr);
				StandingsRequestCreatorConfig requestConfig = new StandingsRequestCreatorConfig
				{
					StartGamesRowNum = startRowNum,
					SheetStartRowIndex = startRowIndex,
					RowCount = poolTeams.Count(),
					SheetId = config.SheetId,
				};
				Request request = requestCreator.CreateRequest(requestConfig);
				requests.Add(request);
			}

			startRowIndex += config.GamesPerRound + 1; // +1 blank row
			info.UpdateValuesRequests.Add(winnerAndPtsHeadersRequest);
			info.UpdateSheetRequests.AddRange(requests);
			return info;
		}
	}
}
