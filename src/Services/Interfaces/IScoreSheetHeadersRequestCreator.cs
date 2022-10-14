namespace PantherShootoutScoreSheetGenerator.Services
{
	public interface IScoreSheetHeadersRequestCreator
	{
		PoolPlayInfo CreateHeaderRequests(PoolPlayInfo info, string poolName, int startRowIndex, IEnumerable<Team> poolTeams);
	}
}