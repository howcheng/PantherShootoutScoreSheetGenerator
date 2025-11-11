using System.ComponentModel;
using System.Text;
using GoogleSheetsHelper;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	/// <summary>
	/// Creates the sorted standings list for a division on the shootout sheet
	/// </summary>
	public sealed class ShootoutSortedStandingsListRequestCreator : SortedStandingsListRequestCreatorBase, IShootoutSortedStandingsListRequestCreator
	{
		private readonly DivisionSheetConfig _config;
		private readonly ShootoutScoringFormulaGenerator _formulaGenerator;
		private readonly ShootoutSheetHelper _helper;

		private static readonly List<Tuple<string, ListSortDirection>> s_sortColumns = new()
		{
			new(Constants.HDR_TOTAL_PTS, ListSortDirection.Descending),
			new(Constants.HDR_TIEBREAKER_GOALS_AGAINST, ListSortDirection.Ascending),
			new(Constants.HDR_TIEBREAKER_KFTM_WINNER, ListSortDirection.Descending),
		};

		public ShootoutSortedStandingsListRequestCreator(FormulaGenerator formulaGenerator, DivisionSheetConfig config)
			: base(formulaGenerator.SheetHelper)
		{
			_formulaGenerator = (ShootoutScoringFormulaGenerator)formulaGenerator;
			_config = config;
			_helper = (ShootoutSheetHelper)_formulaGenerator.SheetHelper;
		}

		public UpdateRequest CreateSortedStandingsListRequest(Tuple<int, int> shootoutStartAndEndRows)
		{
			// =SORT({A3:A12,F3:F12,AA3:AA12,AB3:AB12},2,false,3,true,4,false)
			// sort array of team names, total points, goals against, and KFTM tiebreaker by
			// 1. points (descending)
			// 2, fewest goals against (ascending)
			// 3. KFTM winner (descending)

			int startRowNum = shootoutStartAndEndRows.Item1;
			int endRowNum = shootoutStartAndEndRows.Item2;

			StringBuilder sb = new("{");
			sb.AppendFormat("{0}{1}:{0}{2}", _helper.TeamNameColumnName, startRowNum, endRowNum);
			foreach (var sortColumn in s_sortColumns)
			{
				string columnName = _helper.GetColumnNameByHeader(sortColumn.Item1);
				sb.AppendFormat(",{0}{1}:{0}{2}", columnName, startRowNum, endRowNum);
			}
			sb.Append('}');

			string combinedRange = sb.ToString();
			UpdateRequest request = CreateSortedStandingsListRequest(ShootoutConstants.SHOOTOUT_SHEET_NAME, combinedRange, startRowNum, s_sortColumns);
			return request;
		}

		protected override int GetSortColumnNumber(string columnHeader)
		{
			Tuple<string, ListSortDirection> match = s_sortColumns.First(pair => pair.Item1 == columnHeader);
			return s_sortColumns.IndexOf(match) + 2; // +2 because skip team name column and not zero-based
		}
	}
}
