namespace PantherShootoutScoreSheetGenerator.Services
{
	public interface IShootoutSheetService
	{
		Task<DivisionSheetConfig> GenerateSheet(IDictionary<string, IEnumerable<Team>> allTeams);
	}
}