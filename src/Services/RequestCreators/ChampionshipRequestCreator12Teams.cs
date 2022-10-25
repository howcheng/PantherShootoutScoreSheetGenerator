using GoogleSheetsHelper;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public class ChampionshipRequestCreator12Teams : ChampionshipRequestCreator
	{
		public ChampionshipRequestCreator12Teams(DivisionSheetConfig config, PsoDivisionSheetHelper helper)
			: base(config, helper)
		{
		}

		/// <summary>
		/// In a 12-team division, there are 3 pools and the championship consists of the two 1st-place teams that have the most points
		/// and the consolation is between the lowest-scoring 1st-place team and the highest-scoring 2nd-place team
		/// </summary>
		/// <param name="info"></param>
		/// <returns></returns>
		public override ChampionshipInfo CreateChampionshipRequests(PoolPlayInfo info)
		{
			ChampionshipInfo ret = new ChampionshipInfo(info);

			int startRowIndex = info.ChampionshipStartRowIndex;
			int startRowNum = startRowIndex + 1;
			List<UpdateRequest> updateRequests = new List<UpdateRequest>();

			// finals label row
			updateRequests.Add(Utilities.CreateHeaderLabelRowRequest(_divisionName, startRowIndex, _helper.GetColumnIndexByHeader(_helper.StandingsTableColumns.Last()), "FINALS", 4));

			// championship and consolation headers and teams
			updateRequests.AddRange(CreateHeaderAndTeamRows(info, false, ref startRowIndex, ref startRowNum, "3RD-PLACE: pool-winner with least pts v 2nd-place team with most pts"));
			ret.ThirdPlaceGameRowNum = startRowNum;
			updateRequests.AddRange(CreateHeaderAndTeamRows(info, true, ref startRowIndex, ref startRowNum, "CHAMPIONSHIP: pool-winners with most pts"));
			ret.ChampionshipGameRowNum = startRowNum;
			return ret;
		}

		private IList<UpdateRequest> CreateHeaderAndTeamRows(PoolPlayInfo poolPlayInfo, bool isChampionship, ref int startRowIndex, ref int startRowNum, string label)
		{
			List<UpdateRequest> updateRequests = new List<UpdateRequest>();
			// header row
			startRowIndex += 1;
			startRowNum += 1;
			string headerCellRange = Utilities.CreateCellRangeString(_helper.HomeTeamColumnName, startRowNum, _helper.AwayTeamColumnName, startRowNum);
			updateRequests.Add(Utilities.CreateHeaderLabelRowRequest(_divisionName, startRowIndex, _helper.GetColumnIndexByHeader(Constants.HDR_AWAY_TEAM), label, 0, cell => cell.SetSubheaderCellFormatting()));

			// team cells
			startRowIndex += 1;
			startRowNum += 1;

			string homeFormula = GetTeamFormula(isChampionship ? 1 : 3, poolPlayInfo);
			string awayFormula = GetTeamFormula(isChampionship ? 2 : 4, poolPlayInfo);
			updateRequests.Add(CreateChampionshipGameRequests(startRowNum, homeFormula, awayFormula));
			return updateRequests;
		}

		private string GetTeamFormula(int rank, PoolPlayInfo inf)
		{
			PoolPlayInfo12Teams info = (PoolPlayInfo12Teams)inf;
			int helperCellsStartRowNum = rank <= 3 ? info.HelperCellStartAndEndRowNums.First().Item1 : info.HelperCellStartAndEndRowNums.Last().Item1;
			int helperCellsEndRowNum = rank <= 3 ? info.HelperCellStartAndEndRowNums.First().Item2 : info.HelperCellStartAndEndRowNums.Last().Item2;
			// =IF(COUNT(W3:W5,"=3")=3, VLOOKUP([rank],{V3:V5,T3:T5},2,FALSE), "")
			// if not all games played yet do nothing
			// otherwise get the name of the team with rank N from the pool-winners or runners-up ranges
			string gamesPlayedCellRange = Utilities.CreateCellRangeString(_helper.GetColumnNameByHeader(ShootoutConstants.HDR_POOL_WINNER_GAMES_PLAYED), helperCellsStartRowNum, helperCellsEndRowNum);
			string rankCellRange = Utilities.CreateCellRangeString(_helper.GetColumnNameByHeader(ShootoutConstants.HDR_POOL_WINNER_RANK), helperCellsStartRowNum, helperCellsEndRowNum);
			string teamNameCellRange = Utilities.CreateCellRangeString(_helper.GetColumnNameByHeader(ShootoutConstants.HDR_POOL_WINNERS), helperCellsStartRowNum, helperCellsEndRowNum);

			int rankForVlookup = rank <= 3 ? rank : 1; // rank 4 = top 2nd-place team
			return $"=IF(COUNTIF({gamesPlayedCellRange}, \"=3\")=3, VLOOKUP({rankForVlookup},{{{rankCellRange},{teamNameCellRange}}},2,FALSE), \"\")";
		}
	}
}
