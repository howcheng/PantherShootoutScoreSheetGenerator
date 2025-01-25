namespace PantherShootoutScoreSheetGenerator.Services
{
	public interface IShootoutSheetService
	{
		Task<ShootoutSheetConfig> GenerateSheet(IDictionary<string, IEnumerable<Team>> allTeams);
	}
}