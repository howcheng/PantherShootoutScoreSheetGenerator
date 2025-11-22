using Google.Apis.Sheets.v4.Data;

namespace PantherShootoutScoreSheetGenerator.Services
{
	/// <summary>
	/// Helper for creating requests to hide columns in Google Sheets
	/// </summary>
	public interface IColumnVisibilityHelper
	{
		/// <summary>
		/// Creates requests to group a range of columns and hide the group
		/// </summary>
		/// <param name="sheetId">The sheet ID</param>
		/// <param name="startColumnIndex">The start column index (0-based, inclusive)</param>
		/// <param name="endColumnIndex">The end column index (0-based, exclusive)</param>
		/// <returns>A list of Request objects that create and hide a column group</returns>
		IList<Request> CreateHideColumnsRequest(int? sheetId, int startColumnIndex, int endColumnIndex);
	}
}
