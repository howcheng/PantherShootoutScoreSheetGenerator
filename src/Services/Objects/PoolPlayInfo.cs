using Google.Apis.Sheets.v4.Data;
using GoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public class SheetRequests
	{
		public IList<Request> UpdateSheetRequests { get; set; } = new List<Request>();
		public IList<UpdateRequest> UpdateValuesRequests { get; set; } = new List<UpdateRequest>();
	}

	/// <summary>
	/// Contains sheet details for a division for building the pool play requests
	/// </summary>
	public class PoolPlayInfo : SheetRequests
	{
		public PoolPlayInfo()
		{
		}

		public PoolPlayInfo(IEnumerable<Team> teams)
			: base()
		{
			Pools = teams.GroupBy(x => x.PoolName).OrderBy(x => x.Key);
		}

		protected PoolPlayInfo(IOrderedEnumerable<IGrouping<string, Team>>? pools)
		{
			Pools = pools;
		}

		public IOrderedEnumerable<IGrouping<string, Team>>? Pools { get; protected set; }

		/// <summary>
		/// The index of the row where the championship info starts; this should be the row with the label header ("FINALS")
		/// </summary>
		public int ChampionshipStartRowIndex { get; set; }
		/// <summary>
		/// The start and end row numbers for each of the standings tables (one for each pool, in order)
		/// </summary>
		public List<Tuple<int, int>> StandingsStartAndEndRowNums { get; set; } = new List<Tuple<int, int>>();
		/// <summary>
		/// The start and end row numbers for the scores entry portion of the sheet (one for each pool, in order); key is the pool name
		/// </summary>
		public Dictionary<string, List<Tuple<int, int>>> ScoreEntryStartAndEndRowNums { get; set; } = new();

		public int FirstScoreEntryRowNum => ScoreEntryStartAndEndRowNums.First().Value.First().Item1;
		public int LastScoreEntryRowNum => ScoreEntryStartAndEndRowNums.Last().Value.Last().Item2;
	}

	public class ChampionshipInfo : PoolPlayInfo
	{
		public int ChampionshipGameRowNum { get; set; }
		public int ThirdPlaceGameRowNum { get; set; }

		public ChampionshipInfo(PoolPlayInfo poolPlayInfo)
			: base()
		{
			Pools = poolPlayInfo.Pools;
			StandingsStartAndEndRowNums = poolPlayInfo.StandingsStartAndEndRowNums;
			ChampionshipStartRowIndex = poolPlayInfo.ChampionshipStartRowIndex;
		}

		public IEnumerable<Tuple<int, int>> StandingsStartAndEndIndices => StandingsStartAndEndRowNums.Select(x => Tuple.Create(x.Item1 - 1, x.Item2 - 1));
		public int FirstStandingsRowNum => StandingsStartAndEndRowNums.First().Item1;
	}
}
