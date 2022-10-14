using GoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public static class ServicesExtensions
	{
		public static GoogleSheetCell SetHeaderCellFormatting(this GoogleSheetCell cell)
			=> cell.SetBackgroundColor(Colors.HeaderRowColor).SetBoldText();

		public static GoogleSheetCell SetSubheaderCellFormatting(this GoogleSheetCell cell)
			=> cell.SetBackgroundColor(Colors.SubheaderRowColor).SetBoldText();

		public static GoogleSheetCell SetBackgroundColor(this GoogleSheetCell cell, System.Drawing.Color color)
			=> cell.SetBackgroundColor(color.ToGoogleColor());

		public static GoogleSheetCell SetBackgroundColor(this GoogleSheetCell cell, Google.Apis.Sheets.v4.Data.Color color)
		{
			cell.GoogleBackgroundColor = color;
			return cell;
		}

		public static GoogleSheetCell SetBoldText(this GoogleSheetCell cell)
		{
			cell.Bold = true;
			return cell;
		}
	}
}
