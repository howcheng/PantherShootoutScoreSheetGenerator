using Google.Apis.Sheets.v4.Data;
using GoogleSheetsHelper;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public class ScoreInputsRequestCreator : IScoreInputsRequestCreator
	{
		private readonly string _divisionName;
		private readonly DivisionSheetConfig _config;
		private readonly PsoDivisionSheetHelper _helper;
		private readonly IStandingsRequestCreatorFactory _requestCreatorFactory;

		public ScoreInputsRequestCreator(DivisionSheetConfig config, PsoDivisionSheetHelper helper, IStandingsRequestCreatorFactory factory)
		{
			_divisionName = config.DivisionName;
			_config = config;
			_helper = helper;
			_requestCreatorFactory = factory;
		}

		public virtual PoolPlayInfo CreateScoringRequests(PoolPlayInfo info, IEnumerable<Team> poolTeams, int roundNum, ref int startRowIndex)
		{
			List<Request> requests = new List<Request>();

			// round label and forfeit header
			int labelRowIndex = startRowIndex;
			UpdateRequest roundHeaderRequest = Utilities.CreateHeaderLabelRowRequest(_divisionName, labelRowIndex, _helper.GetColumnIndexByHeader(Constants.HDR_AWAY_TEAM) + 1, $"ROUND {roundNum}", 0, cell => cell.SetSubheaderCellFormatting());
			roundHeaderRequest.Rows.Single().Last().StringValue = Constants.HDR_FORFEIT;
			info.UpdateValuesRequests.Add(roundHeaderRequest);

			// scoring rows: home team name, away team name, checkbox for forfeit
			int gameStartRowIndex = labelRowIndex + 1;
			int firstGameRowNum = gameStartRowIndex + 1;
			string firstTeamSheetCell = poolTeams.First().TeamSheetCell;
			string lastTeamSheetCell = poolTeams.Last().TeamSheetCell;
			requests.Add(RequestCreator.CreateDataValidationRequest(ShootoutConstants.SHOOTOUT_SHEET_NAME, firstTeamSheetCell, lastTeamSheetCell, _config.SheetId, gameStartRowIndex, _helper.GetColumnIndexByHeader(Constants.HDR_HOME_TEAM), _config.GamesPerRound));
			requests.Add(RequestCreator.CreateDataValidationRequest(ShootoutConstants.SHOOTOUT_SHEET_NAME, firstTeamSheetCell, lastTeamSheetCell, _config.SheetId, gameStartRowIndex, _helper.GetColumnIndexByHeader(Constants.HDR_AWAY_TEAM), _config.GamesPerRound));
			IStandingsRequestCreator forfeitRequestCreator = _requestCreatorFactory.GetRequestCreator(Constants.HDR_FORFEIT);
			StandingsRequestCreatorConfig requestConfig = new StandingsRequestCreatorConfig
			{
				SheetId = _config.SheetId,
				SheetStartRowIndex = gameStartRowIndex,
				RowCount = _config.GamesPerRound,
				StartGamesRowNum = firstGameRowNum
			};
			Request forfeitRequest = forfeitRequestCreator.CreateRequest(requestConfig);
			requests.Add(forfeitRequest);

			UpdateRequest winnerAndPtsHeadersRequest = new UpdateRequest(_divisionName)
			{
				RowStart = labelRowIndex,
				ColumnStart = _helper.GetColumnIndexByHeader(PsoDivisionSheetHelper.WinnerAndPointsColumns.First()),
			};
			winnerAndPtsHeadersRequest.Rows.Add(_helper.CreateHeaderRow(PsoDivisionSheetHelper.WinnerAndPointsColumns));

			foreach (string hdr in PsoDivisionSheetHelper.WinnerAndPointsColumns)
			{
				IStandingsRequestCreator requestCreator = _requestCreatorFactory.GetRequestCreator(hdr);
				Request request = requestCreator.CreateRequest(requestConfig);
				requests.Add(request);
			}

			startRowIndex = gameStartRowIndex + _config.GamesPerRound + 1;
			info.UpdateValuesRequests.Add(winnerAndPtsHeadersRequest);
			info.UpdateSheetRequests.AddRange(requests);
			info.UpdateSheetRequests.AddRange(_helper.CreateCellWidthRequests(_config.SheetId, _helper.StandingsTableColumns, _config.TeamNameCellWidth));
			return info;
		}
	}
}
