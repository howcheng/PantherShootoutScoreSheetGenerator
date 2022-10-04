using GoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public static class ServicesExtensions
	{
		public static GoogleSheetCell SetHeaderCellFormatting(this GoogleSheetCell cell)
		{
			cell.BackgroundColor = Colors.HeaderRowColor;
			cell.Bold = true;
			return cell;
		}

		public static GoogleSheetCell SetSubheaderCellFormatting(this GoogleSheetCell cell)
		{
			cell.BackgroundColor = Colors.SubheaderRowColor;
			cell.Bold = true;
			return cell;
		}

	}
}
