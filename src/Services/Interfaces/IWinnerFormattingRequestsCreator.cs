namespace PantherShootoutScoreSheetGenerator.Services
{
	public interface IWinnerFormattingRequestsCreator
	{
		SheetRequests CreateWinnerFormattingRequests(ChampionshipInfo info);
	}
}