using System.ComponentModel;
using System.Text;
using GoogleSheetsHelper;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	/// <summary>
	/// Creates the pool-winners sorted standings lists for a 12-team division, which applies the tiebreakers to the standings list
	/// </summary>
	public sealed class PoolWinnersSortedStandingsListRequestCreator : SortedStandingsListRequestCreator, IPoolWinnersSortedStandingsListRequestCreator
	{
		private readonly List<Tuple<string, ListSortDirection>> _sortColumns;
		private readonly DivisionSheetConfig _config;

		public PoolWinnersSortedStandingsListRequestCreator(FormulaGenerator formulaGenerator, DivisionSheetConfig config) 
			: base(formulaGenerator, config)
		{
			// Create a new list from the base sort columns instead of referencing the static list
			_sortColumns = new List<Tuple<string, ListSortDirection>>(s_sortColumns);
			// we aren't using the head-to-head tiebreaker
			// and although we are not using the KFTM tiebreaker, we're replacing it with the pool-winners tiebreaker box, so it's in the same place
			_sortColumns.Remove(_sortColumns.Single(x => x.Item1 == Constants.HDR_TIEBREAKER_H2H));
			_config = config;
		}

		public override PoolPlayInfo CreateSortedStandingsListRequest(PoolPlayInfo info)
		{
			info = CreateSortedStandingsListRequest(info, true);
			info = CreateSortedStandingsListRequest(info, false);
			return info;
		}

		private PoolPlayInfo CreateSortedStandingsListRequest(PoolPlayInfo info, bool isPoolWinners)
		{
			// =SORT({{AI3:BB3,AF42};{AI16:BB16,AF43};{AI29:BB29,AF44}},8,false,17,false,18,true,19,true,20,false,21,true)
			// note that we are using the other sorted standings lists -- the 1st/2nd place teams
			// the sort array is the regular standings with tiebreakers plus the pool-winners tiebreaker checkbox column
			// we also are ignoring the head-to-head tiebreaker
			StringBuilder sb = new("{");
			bool first = true;
			int startRowNum = info.ChampionshipStartRowIndex + 3; // the index will be for the championship header row
			if (!isPoolWinners)
				startRowNum += _helper.DivisionSheetConfig.TeamsPerPool;
			int endRowNum = startRowNum + _helper.DivisionSheetConfig.TeamsPerPool - 1; // 1 line per pool

			int startColIdx = _helper.SortedStandingsListColumnIndex;
			string startColName = Utilities.ConvertIndexToColumnName(startColIdx);
			// because the sorted standings list starts with the team name, the end column index 
			int endColIdx = startColIdx + GetSortColumnNumber(Constants.HDR_TIEBREAKER_GOAL_DIFF) - 1; // converting back to index
			string endColName = Utilities.ConvertIndexToColumnName(endColIdx);
			string tbCbxColName = _helper.GetColumnNameByHeader(ShootoutConstants.HDR_POOL_WINNER_TIEBREAKER);

			// loop thru the standings entries to get the 1st/2nd place winner row
			List<string> cellRanges = new();
			foreach (var startAndEnd in info.StandingsStartAndEndRowNums)
			{
				int rowNum = startAndEnd.Item1;
				if (!isPoolWinners)
					rowNum += 1;
				string cellRange = Utilities.CreateCellRangeString(startColName, rowNum, endColName, rowNum);
				cellRanges.Add(cellRange);
			}

			// loop thru the pool winner rows to make the tiebreaker checkbox cell ref
			for (int i = 0; i < _helper.DivisionSheetConfig.NumberOfPools; i++)
			{
				int rowNum = startRowNum + i;
				string cellRange = cellRanges[i];
				string tbCbxCell = Utilities.CreateCellReference(tbCbxColName, rowNum);

				if (!first)
					sb.Append(';');
				sb.AppendFormat("{{{0},{1}}}", cellRange, tbCbxCell);
				first = false;
			}
			sb.Append('}');

			string combinedRange = sb.ToString();
			UpdateRequest request = CreateSortedStandingsListRequest(_config.DivisionName, combinedRange, startRowNum, _sortColumns);
			info.UpdateValuesRequests.Add(request);
			return info;
		}
	}
}
