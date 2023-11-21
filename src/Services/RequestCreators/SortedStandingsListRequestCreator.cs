using GoogleSheetsHelper;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public interface ISortedStandingsListRequestCreator
	{
		PoolPlayInfo CreateSortedStandingsListRequest(PoolPlayInfo info, int startRowNum, int endRowNum);
	}

	public class SortedStandingsListRequestCreator : ISortedStandingsListRequestCreator
	{
		private readonly DivisionSheetConfig _config;
		private readonly PsoFormulaGenerator _formulaGenerator;
		private readonly PsoDivisionSheetHelper _helper;

		public SortedStandingsListRequestCreator(FormulaGenerator formulaGenerator, DivisionSheetConfig config)
		{
			_formulaGenerator = (PsoFormulaGenerator)formulaGenerator;
			_helper = (PsoDivisionSheetHelper)_formulaGenerator.SheetHelper;
			_config = config;
		}

		public PoolPlayInfo CreateSortedStandingsListRequest(PoolPlayInfo info, int startRowNum, int endRowNum)
		{
			// =SORT(G3:AA5,15,true,16,false,17,true,18,false,19,true,20,false,21,false) 15 = the relative cell number (counting from G3) containing the calculated rank formula
			string cellRange = Utilities.CreateCellRangeString(_helper.TeamNameColumnName, startRowNum, _helper.KicksFromTheMarkTiebreakerColumnName, endRowNum);
			int startColNum = _helper.GetColumnIndexByHeader(Constants.HDR_CALC_RANK) - _helper.GetColumnIndexByHeader(Constants.HDR_TEAM_NAME) + 1;
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
							FormulaValue = $"=SORT({cellRange},{startColNum++},true,{startColNum++},false,{startColNum++},true,{startColNum++},false,{startColNum++},true,{startColNum++},false,{startColNum++},false)",
						}
					}
				},
			};
			info.UpdateValuesRequests.Add(request);
			return info;
		}
	}
}
