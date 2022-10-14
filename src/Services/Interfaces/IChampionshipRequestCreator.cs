namespace PantherShootoutScoreSheetGenerator.Services
{
	public interface IChampionshipRequestCreator
	{
		ChampionshipInfo CreateChampionshipRequests(PoolPlayInfo poolPlayInfo);
	}
}
