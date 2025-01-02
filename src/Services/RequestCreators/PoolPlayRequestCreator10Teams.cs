using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public sealed class PoolPlayRequestCreator10Teams : PoolPlayRequestCreator
	{
		private readonly IStandingsRequestCreatorFactory _requestCreatorFactory;
		private readonly ISortedStandingsListRequestCreator _sortedStandingsListCreator;
		private static readonly ISortedStandingsListRequestCreator s_mockSortedStandingsCreator = new MockSortedStandingsListRequestCreator();

		public PoolPlayRequestCreator10Teams(DivisionSheetConfig config, IScoreSheetHeadersRequestCreator headersCreator, IScoreInputsRequestCreator inputsCreator, IStandingsTableRequestCreator tableCreator
			, IStandingsRequestCreatorFactory factory, ITiebreakerColumnsRequestCreator tiebreakersCreator, ISortedStandingsListRequestCreator sortedStandingsCreator) 
			: base(config, headersCreator, inputsCreator, tableCreator, tiebreakersCreator, s_mockSortedStandingsCreator)
		{
			_requestCreatorFactory = factory;
			_sortedStandingsListCreator = sortedStandingsCreator;
		}

		public override PoolPlayInfo CreatePoolPlayRequests(PoolPlayInfo info)
		{
			info = base.CreatePoolPlayRequests(info);

			// we had to wait for all the pools to be built before we can do the sorted standings list
			info = _sortedStandingsListCreator.CreateSortedStandingsListRequest(info);

			return info;
		}

		// we have to wait until the end to create the sorted standings list, so this is a dummy class that does nothing 
		private class MockSortedStandingsListRequestCreator : ISortedStandingsListRequestCreator
		{
			public PoolPlayInfo CreateSortedStandingsListRequest(PoolPlayInfo info)
				=> info;
		}
	}
}
