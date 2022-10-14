using Google.Apis.Sheets.v4.Data;
using GoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public class SheetRequests
	{
		public IList<Request> UpdateSheetRequests { get; set; } = new List<Request>();
		public IList<UpdateRequest> UpdateValuesRequests { get; set; } = new List<UpdateRequest>();
	}

	public class PoolPlayInfo : SheetRequests
	{
		/// <summary>
		/// This constructor is required for AutoFixture because it doesn't know how to autogenerate a IOrderedEnumerable
		/// </summary>
		public PoolPlayInfo()
		{
		}

		public PoolPlayInfo(IOrderedEnumerable<IGrouping<string, Team>>? pools)
			: base()
		{
			Pools = pools;
		}

		public IOrderedEnumerable<IGrouping<string, Team>>? Pools { get; set; }
		public int TeamsPerPool => Pools?.FirstOrDefault()?.Count() ?? 0;

		public int ChampionshipStartRowIndex { get; set; }
		/// <summary>
		/// The start and end row numbers for each of the standings tables (one for each pool, in order)
		/// </summary>
		public List<Tuple<int, int>> StandingsStartAndEndRowNums { get; set; } = new List<Tuple<int, int>>();
	}

	public class ChampionshipInfo : PoolPlayInfo
	{
		public int ChampionshipGameRowNum { get; set; }
		public int ThirdPlaceGameRowNum { get; set; }

		public ChampionshipInfo(PoolPlayInfo poolPlayInfo)
			: base(poolPlayInfo.Pools)
		{
			StandingsStartAndEndRowNums = poolPlayInfo.StandingsStartAndEndRowNums;
			ChampionshipStartRowIndex = poolPlayInfo.ChampionshipStartRowIndex;
		}

		public IEnumerable<Tuple<int, int>> StandingsStartAndEndIndices => StandingsStartAndEndRowNums.Select(x => Tuple.Create(x.Item1 - 1, x.Item2 - 1));
		public int FirstStandingsRowNum => StandingsStartAndEndRowNums.First().Item1;
	}
}
