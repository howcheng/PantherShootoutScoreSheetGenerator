using Google.Apis.Sheets.v4.Data;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public interface IDivisionSheetGenerator
	{
		Task<PoolPlayInfo> CreateSheet(ShootoutSheetConfig shootoutSheetConfig);
	}
}
