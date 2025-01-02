using System.Text;
using GoogleSheetsHelper;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	/// <summary>
	/// Creates the sorted standings list for a 10-team division, which applies the tiebreakers to the standings list
	/// </summary>
	public sealed class SortedStandingsListRequestCreator10Teams : SortedStandingsListRequestCreator
	{
		public SortedStandingsListRequestCreator10Teams(FormulaGenerator formulaGenerator, DivisionSheetConfig config)
			: base(formulaGenerator, config)
		{
		}

		public override PoolPlayInfo CreateSortedStandingsListRequest(PoolPlayInfo info)
		{
			// =SORT({G3:AA5;G24:G29},8,false,15,true,16,false,17,true,18,false,19,true,20,false,21,false) 15 = the relative cell number (counting from G3) containing the calculated rank formula
			StringBuilder sb = new("{");
			bool first = true;
			foreach (var startAndEnd in info.StandingsStartAndEndRowNums)
			{
				int startRowNum = startAndEnd.Item1;
				int endRowNum = startAndEnd.Item2;
				string cellRange = Utilities.CreateCellRangeString(_helper.TeamNameColumnName, startRowNum, _helper.KicksFromTheMarkTiebreakerColumnName, endRowNum);
				if (!first)
					sb.Append(';');
				sb.Append(cellRange);
				first = false;
			}
			sb.Append('}');

			string combinedRange = sb.ToString();
			return CreateSortedStandingsListRequest(info, combinedRange, info.StandingsStartAndEndRowNums.First().Item1);
		}
	}
}
