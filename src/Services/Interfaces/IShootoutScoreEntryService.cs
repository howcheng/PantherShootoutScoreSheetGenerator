
namespace PantherShootoutScoreSheetGenerator.Services
{
	public interface IShootoutScoreEntryService
	{
		Task CreateShootoutScoreEntrySection(PoolPlayInfo info);
	}
}