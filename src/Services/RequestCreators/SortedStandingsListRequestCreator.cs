using System.ComponentModel;
using System.Text;
using GoogleSheetsHelper;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	/// <summary>
	/// Creates the sorted standings list, which applies the tiebreakers to the standings list
	/// </summary>
	public class SortedStandingsListRequestCreator : ISortedStandingsListRequestCreator
	{
		private readonly DivisionSheetConfig _config;
		private readonly PsoFormulaGenerator _formulaGenerator;
		protected readonly PsoDivisionSheetHelper _helper;

		protected static readonly List<Tuple<string, ListSortDirection>> s_sortColumns = new()
		{
			new(Constants.HDR_GAME_PTS, ListSortDirection.Descending),
			new(Constants.HDR_TIEBREAKER_H2H, ListSortDirection.Descending),
			new(Constants.HDR_TIEBREAKER_WINS, ListSortDirection.Descending),
			new(Constants.HDR_TIEBREAKER_CARDS, ListSortDirection.Ascending),
			new(Constants.HDR_TIEBREAKER_GOALS_AGAINST, ListSortDirection.Ascending),
			new(Constants.HDR_TIEBREAKER_GOAL_DIFF, ListSortDirection.Descending),
			new(Constants.HDR_TIEBREAKER_KFTM_WINNER, ListSortDirection.Descending),
		};

		public SortedStandingsListRequestCreator(FormulaGenerator formulaGenerator, DivisionSheetConfig config)
		{
			_formulaGenerator = (PsoFormulaGenerator)formulaGenerator;
			_helper = (PsoDivisionSheetHelper)_formulaGenerator.SheetHelper;
			_config = config;
		}

		public virtual PoolPlayInfo CreateSortedStandingsListRequest(PoolPlayInfo info)
		{
			// =SORT(G3:AA5,15,true,16,false,17,true,18,false,19,true,20,false,21,false) 15 = the relative cell number (counting from G3) containing the calculated rank formula
			Tuple<int, int> standingsStartAndEnd = info.StandingsStartAndEndRowNums.Last(); // because we are creating one per pool, the latest one is what we want
			int startRowNum = standingsStartAndEnd.Item1;
			int endRowNum = standingsStartAndEnd.Item2;
			string cellRange = Utilities.CreateCellRangeString(_helper.TeamNameColumnName, startRowNum, _helper.KicksFromTheMarkTiebreakerColumnName, endRowNum);
			return CreateSortedStandingsListRequest(info, cellRange, startRowNum);
		}

		protected PoolPlayInfo CreateSortedStandingsListRequest(PoolPlayInfo info, string cellRange, int startRowNum)
		{
			// for Boolean columns (H2H, KTFM), FALSE comes before TRUE, so those should be sorted descending
			// sort rules:
			// 1. most game points (descending)
			// 2. head-to-head tiebreaker (descending)
			// 3. most wins (descending)
			// 4. fewest cards (ascending)
			// 5. fewest goals against (ascending)
			// 6. biggest goal differential (descending)
			// 7. KFTM winner (descending)
			return CreateSortedStandingsListRequest(info, cellRange, startRowNum, s_sortColumns);
		}

		protected PoolPlayInfo CreateSortedStandingsListRequest(PoolPlayInfo info, string cellRange, int startRowNum, IEnumerable<Tuple<string, ListSortDirection>> sortColumns)
		{
			// column numbers for sorting are relative to the array, not the actual column numbers or indexes
			// (i.e., if G3 is the first column in the array, then column G will be 1)

			StringBuilder sb = new("=SORT(");
			sb.Append(cellRange);
			sb.Append(',');
			bool first = true;
			foreach (var item in sortColumns)
			{
				string col = item.Item1;
				ListSortDirection dir = item.Item2;

				int colNum = GetSortColumnNumber(col);
				string sort = Utilities.CreateSortColumnReference(colNum, dir);
				if (!first)
					sb.Append(',');
				sb.Append(sort);
				first = false;
			}
			sb.Append(')');

			UpdateRequest request = new(_config.DivisionName)
			{
				ColumnStart = _helper.SortedStandingsListColumnIndex,
				RowStart = startRowNum - 1,
				Rows = new List<GoogleSheetRow>
				{
					new GoogleSheetRow
					{
						new GoogleSheetCell
						{
							FormulaValue = sb.ToString(),
						}
					}
				},
			};
			info.UpdateValuesRequests.Add(request);
			return info;
		}

		protected int GetSortColumnNumber(string columnHeader)
		{
			int teamNameColNum = _helper.GetColumnIndexByHeader(Constants.HDR_TEAM_NAME);
			int colNum = _helper.GetColumnIndexByHeader(columnHeader);
			return colNum - teamNameColNum + 1; // since it's not zero-based
		}
	}
}
