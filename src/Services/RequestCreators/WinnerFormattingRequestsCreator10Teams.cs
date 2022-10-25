using Google.Apis.Sheets.v4.Data;
using GoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public class WinnerFormattingRequestsCreator10Teams : WinnerFormattingRequestsCreator
	{
		public WinnerFormattingRequestsCreator10Teams(DivisionSheetConfig config, PsoDivisionSheetHelper helper)
			: base(config, helper)
		{
		}

		protected override Request CreateWinnerFormattingRequest(int rank, ChampionshipInfo championshipInfo)
		{
			// =AND($M3=1,COUNTIF($N$3:$N$7,"=4")=5,COUNTIF($N$24:$N$28,"=4")=5)
			List<Tuple<int, int>> standingsStartAndEndRowNums = championshipInfo.StandingsStartAndEndRowNums;
			string rankCell = $"${_helper.GetColumnNameByHeader(ShootoutConstants.HDR_OVERALL_RANK)}{championshipInfo.FirstStandingsRowNum}";
			string pool1Range = Utilities.CreateCellRangeString(_helper.GamesPlayedColumnName, standingsStartAndEndRowNums.First().Item1, standingsStartAndEndRowNums.First().Item2, CellRangeOptions.FixColumn | CellRangeOptions.FixRow);
			string pool2Range = Utilities.CreateCellRangeString(_helper.GamesPlayedColumnName, standingsStartAndEndRowNums.Last().Item1, standingsStartAndEndRowNums.Last().Item2, CellRangeOptions.FixColumn | CellRangeOptions.FixRow);
			string formula = $"=AND({rankCell}=1,COUNTIF({pool1Range},\"=4\")=5,COUNTIF({pool2Range},\"=4\")=5)";

			return CreateWinnerConditionalFormattingRequest(rank, formula, championshipInfo.StandingsStartAndEndIndices);
		}
	}
}
