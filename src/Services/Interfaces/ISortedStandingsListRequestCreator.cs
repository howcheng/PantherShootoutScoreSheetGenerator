using GoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public interface ISortedStandingsListRequestCreator
	{
		/// <summary>
		/// Creates the Google requests to build the sorted standings list
		/// </summary>
		/// <param name="info"></param>
		/// <returns></returns>
		PoolPlayInfo CreateSortedStandingsListRequest(PoolPlayInfo info);
	}

	public interface IPoolWinnersSortedStandingsListRequestCreator : ISortedStandingsListRequestCreator
	{
	}

	public interface IShootoutSortedStandingsListRequestCreator
	{
		UpdateRequest CreateSortedStandingsListRequest(Tuple<int, int> shootoutStartAndEndRows);
	}
}
