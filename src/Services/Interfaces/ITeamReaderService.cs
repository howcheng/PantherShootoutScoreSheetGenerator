namespace PantherShootoutScoreSheetGenerator.Services
{
	/// <summary>
	/// Interface for a service class to load team data
	/// </summary>
	public interface ITeamReaderService
	{
		/// <summary>
		/// Dictionary of teams, where the key is the division name
		/// </summary>
		/// <returns></returns>
		IDictionary<string, IEnumerable<Team>> GetTeams();
	}
}