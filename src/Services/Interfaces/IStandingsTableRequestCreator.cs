namespace PantherShootoutScoreSheetGenerator.Services
{
	public interface IStandingsTableRequestCreator
	{
		/// <summary>
		/// Creates the Google requests to build the standings table
		/// </summary>
		/// <param name="info"></param>
		/// <param name="poolTeams"></param>
		/// <param name="startRowIndex">The index of the first row in the table</param>
		/// <returns></returns>
		PoolPlayInfo CreateStandingsRequests(PoolPlayInfo info, IEnumerable<Team> poolTeams, int startRowIndex);
	}
}