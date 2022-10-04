namespace PantherShootoutScoreSheetGenerator.Services
{
	public interface IShootoutSheetService
	{
		Task GenerateSheet(IDictionary<string, IEnumerable<Team>> allTeams);
	}
}