namespace PantherShootoutScoreSheetGenerator.Services
{
	public interface ITiebreakerColumnsRequestCreator
	{
		PoolPlayInfo CreateTiebreakerRequests(PoolPlayInfo info, IEnumerable<Team> poolTeams, int startRowIndex);
	}
}
