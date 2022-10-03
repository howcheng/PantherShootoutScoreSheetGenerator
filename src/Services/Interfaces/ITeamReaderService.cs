namespace PantherShootoutScoreSheetGenerator.Services
{
	public interface ITeamReaderService
	{
		IDictionary<string, IEnumerable<Team>> GetTeams();
	}
}