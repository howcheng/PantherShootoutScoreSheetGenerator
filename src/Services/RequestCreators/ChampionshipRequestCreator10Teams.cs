using Google.Apis.Sheets.v4.Data;
using GoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public class ChampionshipRequestCreator10Teams : ChampionshipRequestCreator
	{
		public ChampionshipRequestCreator10Teams(string divisionName, DivisionSheetConfig config, PsoDivisionSheetHelper helper)
			: base(divisionName, config, helper)
		{
		}

		/// <summary>
		/// In a 10-team division (2 pools of 5 teams each), there are no championship rounds; winners are determined by total points in pool play across both pools
		/// </summary>
		/// <param name="poolPlayInfo"></param>
		/// <returns></returns>
		public override ChampionshipInfo CreateChampionshipRequests(PoolPlayInfo poolPlayInfo) => new ChampionshipInfo(poolPlayInfo);
	}

	public class WinnerFormattingRequestsCreator10Teams : WinnerFormattingRequestsCreator
	{
		public WinnerFormattingRequestsCreator10Teams(string divisionName, PsoDivisionSheetHelper helper)
			: base(divisionName, helper)
		{
		}

		protected override Request CreateWinnerFormattingRequest(DivisionSheetConfig config, int rank, ChampionshipInfo championshipInfo)
		{
			// =AND($M3=1,COUNTIF($N$3:$N$7,"=4")=5,COUNTIF($N$24:$N$28,"=4")=5)
			List<Tuple<int, int>> standingsStartAndEndRowNums = championshipInfo.StandingsStartAndEndRowNums;
			string rankCell = $"${_helper.GetColumnNameByHeader(ShootoutConstants.HDR_OVERALL_RANK)}{championshipInfo.FirstStandingsRowNum}";
			string pool1Range = Utilities.CreateCellRangeString(_helper.GamesPlayedColumnName, standingsStartAndEndRowNums.First().Item1, standingsStartAndEndRowNums.First().Item2, CellRangeOptions.FixColumn | CellRangeOptions.FixRow);
			string pool2Range = Utilities.CreateCellRangeString(_helper.GamesPlayedColumnName, standingsStartAndEndRowNums.Last().Item1, standingsStartAndEndRowNums.Last().Item2, CellRangeOptions.FixColumn | CellRangeOptions.FixRow);
			string formula = $"=AND({rankCell}=1,COUNTIF({pool1Range},\"=4\")=5,COUNTIF({pool2Range},\"=4\")=5)";

			return CreateWinnerConditionalFormattingRequest(config, rank, formula, championshipInfo.StandingsStartAndEndIndices);
		}
	}
}
