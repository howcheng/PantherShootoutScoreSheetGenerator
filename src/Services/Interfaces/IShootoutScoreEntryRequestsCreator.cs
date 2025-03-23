namespace PantherShootoutScoreSheetGenerator.Services
{
	public interface IShootoutScoreEntryRequestsCreator
	{
		SheetRequests CreateScoreEntryRequests(PoolPlayInfo info, int startRowIndex);
	}
}