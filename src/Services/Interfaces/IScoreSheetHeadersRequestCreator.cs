namespace PantherShootoutScoreSheetGenerator.Services
{
	public interface IScoreSheetHeadersRequestCreator
	{
		/// <summary>
		/// Creates the Google requests to create the header row for the score sheet: standings and head-to-head tiebreakers
		/// </summary>
		/// <param name="info"></param>
		/// <param name="poolName">Name of the pool (e.g. "A", "B")</param>
		/// <param name="rowIndex">The index of the row where the headers will go</param>
		/// <param name="poolTeams">The teams in the pool (needed for the tiebreaker columns)</param>
		/// <returns></returns>
		PoolPlayInfo CreateHeaderRequests(PoolPlayInfo info, string poolName, int rowIndex, IEnumerable<Team> poolTeams);
	}
}