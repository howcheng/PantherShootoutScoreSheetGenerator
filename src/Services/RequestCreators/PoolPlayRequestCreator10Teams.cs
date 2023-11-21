using Google.Apis.Sheets.v4.Data;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public class PoolPlayRequestCreator10Teams : PoolPlayRequestCreator
	{
		private readonly IStandingsRequestCreatorFactory _requestCreatorFactory;

		public PoolPlayRequestCreator10Teams(DivisionSheetConfig config, IScoreSheetHeadersRequestCreator headersCreator, IScoreInputsRequestCreator inputsCreator, IStandingsTableRequestCreator tableCreator
			, IStandingsRequestCreatorFactory factory, ITiebreakerColumnsRequestCreator tiebreakersCreator, ISortedStandingsListRequestCreator sortedStandingsCreator) 
			: base(config, headersCreator, inputsCreator, tableCreator, tiebreakersCreator, sortedStandingsCreator)
		{
			_requestCreatorFactory = factory;
		}

		public override PoolPlayInfo CreatePoolPlayRequests(PoolPlayInfo info)
		{
			info = base.CreatePoolPlayRequests(info);

			// add the request for overall rank in the standings table, which has to be done after the standings tables are already created
			IStandingsRequestCreator rankWithTbCreator = _requestCreatorFactory.Creators.Single(x => x.ColumnHeader == Constants.HDR_RANK);
			IStandingsRequestCreator overallRankCreator = _requestCreatorFactory.Creators.Single(x => x.ColumnHeader == ShootoutConstants.HDR_OVERALL_RANK);
			foreach (var startAndEnd in info.StandingsStartAndEndRowNums)
			{
				int startRowNum = startAndEnd.Item1;
				RankRequestCreatorConfig10Teams config = new RankRequestCreatorConfig10Teams
				{
					SheetId = _config.SheetId,
					SheetStartRowIndex = startRowNum - 1,
					RowCount = _config.TeamsPerPool,
					StartGamesRowNum = startRowNum,
					StandingsStartAndEndRowNums = info.StandingsStartAndEndRowNums,
				};
				Request request1 = rankWithTbCreator.CreateRequest(config);
				info.UpdateSheetRequests.Add(request1);
				
				Request request2 = overallRankCreator.CreateRequest(config);
				info.UpdateSheetRequests.Add(request2);
			}

			return info;
		}
	}
}
