using Google.Apis.Sheets.v4.Data;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public class PoolPlayRequestCreator10Teams : PoolPlayRequestCreator
	{
		private readonly IStandingsRequestCreatorFactory _requestCreatorFactory;

		public PoolPlayRequestCreator10Teams(DivisionSheetConfig config, IScoreSheetHeadersRequestCreator headersCreator, IScoreInputsRequestCreator inputsCreator, IStandingsTableRequestCreator tableCreator
			, IStandingsRequestCreatorFactory factory) 
			: base(config, headersCreator, inputsCreator, tableCreator)
		{
			_requestCreatorFactory = factory;
		}

		public override PoolPlayInfo CreatePoolPlayRequests(PoolPlayInfo info)
		{
			info = base.CreatePoolPlayRequests(info);

			// add the request for overall rank in the standings table, which has to be done after the standings tables are already created
			IStandingsRequestCreator overallRankCreator = _requestCreatorFactory.Creators.Single(x => x.ColumnHeader == ShootoutConstants.HDR_OVERALL_RANK);
			int startRowNum = info.StandingsStartAndEndRowNums.First().Item1;
			OverallRankRequestCreatorConfig config = new OverallRankRequestCreatorConfig
			{
				SheetId = _config.SheetId,
				SheetStartRowIndex = startRowNum - 1,
				RowCount = _config.TeamsPerPool,
				StartGamesRowNum = startRowNum,
				StandingsStartAndEndRowNums = info.StandingsStartAndEndRowNums,
			};
			Request request = overallRankCreator.CreateRequest(config);
			info.UpdateSheetRequests.Add(request);

			return info;
		}
	}
}
