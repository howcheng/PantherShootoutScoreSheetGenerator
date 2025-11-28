namespace PantherShootoutScoreSheetGenerator.Services
{
	public interface IShootoutSheetService
	{
		Task<ShootoutSheetConfig> GenerateSheet(IDictionary<string, IEnumerable<Team>> allTeams, IDictionary<string, DivisionSheetConfig> divisionConfigs);
		Task HideHelperColumns(ShootoutSheetConfig config);
	}
}