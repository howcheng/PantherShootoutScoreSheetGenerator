using Google.Apis.Sheets.v4.Data;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public class WinnerFormattingRequestsCreator10Teams : WinnerFormattingRequestsCreator
	{
		public WinnerFormattingRequestsCreator10Teams(DivisionSheetConfig config, FormulaGenerator fg)
			: base(config, fg)
		{
		}

		protected override Request CreateWinnerFormattingRequest(int rank, ChampionshipInfo championshipInfo)
		{
			string formula = _formulaGenerator.GetConditionalFormattingForWinnerFormula10Teams(rank, championshipInfo.StandingsStartAndEndRowNums, championshipInfo.FirstStandingsRowNum + 1);
			return CreateWinnerConditionalFormattingRequest(rank, formula, championshipInfo.StandingsStartAndEndIndices);
		}
	}
}
