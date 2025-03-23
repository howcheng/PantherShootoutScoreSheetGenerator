namespace PantherShootoutScoreSheetGenerator.Services
{
	/// <summary>
	/// Interface for PSO versions of the <see cref="GoogleSheetsHelper.SheetHelper"/>
	/// </summary>
	public interface IPsoSheetHelper
	{
		DivisionSheetConfig DivisionSheetConfig { get; }
		int SortedStandingsListColumnIndex { get; }
	}
}
