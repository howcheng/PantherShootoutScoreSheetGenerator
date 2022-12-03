using Google.Apis.Sheets.v4.Data;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public class WinnerFormattingRequestsCreator5Teams : WinnerFormattingRequestsCreator
	{
		public WinnerFormattingRequestsCreator5Teams(DivisionSheetConfig config, FormulaGenerator fg)
			: base(config, fg)
		{
		}

		protected override Request CreateWinnerFormattingRequest(int rank, ChampionshipInfo championshipInfo)
		{
			string formula = _formulaGenerator.GetConditionalFormattingForWinnerFormula5Teams(rank, championshipInfo.StandingsStartAndEndRowNums, championshipInfo.FirstStandingsRowNum);
			return CreateWinnerConditionalFormattingRequest(rank, formula, championshipInfo.StandingsStartAndEndIndices);
		}
	}
}
