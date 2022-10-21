namespace PantherShootoutScoreSheetGenerator.Services
{
	public class PoolPlayInfo12Teams : PoolPlayInfo
	{
		public PoolPlayInfo12Teams(IEnumerable<Team> teams) : base(teams)
		{
		}

		public PoolPlayInfo12Teams(PoolPlayInfo poolPlayInfo)
			: base(poolPlayInfo.Pools)
		{
			UpdateSheetRequests = poolPlayInfo.UpdateSheetRequests;
			UpdateValuesRequests = poolPlayInfo.UpdateValuesRequests;
			ChampionshipStartRowIndex = poolPlayInfo.ChampionshipStartRowIndex;
			StandingsStartAndEndRowNums = poolPlayInfo.StandingsStartAndEndRowNums;
		}

		/// <summary>
		/// Holds the start and end row numbers for the pool winners and runners up sections. Item1 is the start row number, and Item2 is the end row number
		/// </summary>
		public List<Tuple<int, int>> HelperCellStartAndEndRowNums { get; set; } = new List<Tuple<int, int>>();
	}
}
