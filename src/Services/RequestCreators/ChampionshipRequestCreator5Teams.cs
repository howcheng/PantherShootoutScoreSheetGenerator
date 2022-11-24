namespace PantherShootoutScoreSheetGenerator.Services
{
	public class ChampionshipRequestCreator5Teams : ChampionshipRequestCreator
	{
		public ChampionshipRequestCreator5Teams(DivisionSheetConfig config, PsoDivisionSheetHelper helper)
			: base(config, helper)
		{
		}

		/// <summary>
		/// In a 10-team division (2 pools of 5 teams each), there are no championship rounds; winners are determined by total points in pool play across both pools
		/// </summary>
		/// <param name="poolPlayInfo"></param>
		/// <returns></returns>
		public override ChampionshipInfo CreateChampionshipRequests(PoolPlayInfo poolPlayInfo) => new ChampionshipInfo(poolPlayInfo);
	}
}
