using Google.Apis.Sheets.v4.Data;
using GoogleSheetsHelper;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public class RankWithTiebreakerRequestCreator : StandingsRequestCreator, IStandingsRequestCreator
	{
		public RankWithTiebreakerRequestCreator(FormulaGenerator formGen) 
			: base(formGen, Constants.HDR_RANK)
		{
		}

		public Request CreateRequest(StandingsRequestCreatorConfig config)
		{
			PsoDivisionSheetHelper helper = (PsoDivisionSheetHelper)_formulaGenerator.SheetHelper;

			// =IFS(OR(O3 >= 3, COUNTIF(O$3:O$6, O3) = 1), O3, NOT(P3), O3+1, P3, O3)
			// if calculated rank >= 3 OR if only one instance of that rank, then do nothing
			// else if the tiebreaker box is not checked, then rank + 1
			// else the rank

			int startRowNum = config.StartGamesRowNum;
			string cellRange = Utilities.CreateCellRangeString(helper.CalculatedRankColumnName, startRowNum, startRowNum + config.RowCount - 1, CellRangeOptions.FixRow);
			string firstRankCell = $"{helper.CalculatedRankColumnName}{startRowNum}";
			string firstTiebreakerCell = $"{helper.TiebreakerColumnName}{startRowNum}";
			string rankWithTbFormula = string.Format("=IFS(OR({0} >= 3, COUNTIF({1}, {0}) = 1), {1}, NOT({2}), {0}+1, {2}, {0})",
				firstRankCell,
				cellRange,
				firstTiebreakerCell);

			Request request = RequestCreator.CreateRepeatedSheetFormulaRequest(config.SheetId, config.SheetStartRowIndex, _columnIndex, config.RowCount,
				rankWithTbFormula);
			return request;
		}
	}

	public class PsoTeamRankRequestCreator : RankRequestCreator
	{
		public PsoTeamRankRequestCreator(FormulaGenerator formGen)
			: base(formGen, Constants.HDR_CALC_RANK)
		{
		}
	}
}
