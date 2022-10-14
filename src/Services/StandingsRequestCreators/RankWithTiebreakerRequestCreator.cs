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

			// =IF(OR(N3 >= 3, COUNTIF(N$3:N$6, N3) = 1), N3, N3 + O3)
			// if calculated rank >= 3 OR if only one instance of that rank, then just it
			// else rank + tiebreaker column

			int startRowNum = config.StartGamesRowNum;
			string cellRange = Utilities.CreateCellRangeString(helper.CalculatedRankColumnName, startRowNum, startRowNum + config.RowCount - 1, CellRangeOptions.FixRow);
			string firstRankCell = $"{helper.CalculatedRankColumnName}{startRowNum}";
			string firstTiebreakerCell = $"{helper.TiebreakerColumnName}{startRowNum}";
			string rankWithTbFormula = string.Format("=IF(OR({0} >= 3, COUNTIF({1}, {0}) = 1), {0}, {0} + {2})",
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
