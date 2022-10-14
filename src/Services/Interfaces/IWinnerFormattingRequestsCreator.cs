namespace PantherShootoutScoreSheetGenerator.Services
{
	public interface IWinnerFormattingRequestsCreator
	{
		SheetRequests CreateWinnerFormattingRequests(DivisionSheetConfig config, ChampionshipInfo info);
	}
}