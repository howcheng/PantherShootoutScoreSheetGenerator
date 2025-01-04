using GoogleSheetsHelper;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public class ChampionshipRequestCreator12Teams : ChampionshipRequestCreator
	{
		private readonly PsoFormulaGenerator _formGen;

		public ChampionshipRequestCreator12Teams(DivisionSheetConfig config, FormulaGenerator formGen)
			: base(config, (PsoDivisionSheetHelper)formGen.SheetHelper)
		{
			_formGen = (PsoFormulaGenerator)formGen;
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
			updateRequests.Add(Utilities.CreateHeaderLabelRowRequest(_divisionName, startRowIndex, _helper.GetColumnIndexByHeader(_helper.StandingsTableColumns.Last()), "FINALS", 4, cell => cell.SetHeaderCellFormatting()));

			// championship and consolation headers and teams
			updateRequests.AddRange(CreateHeaderAndTeamRows(info, false, ref startRowIndex, ref startRowNum, "3RD-PLACE: 3rd-place pool-winner v 1st-place pool runner-up"));
			ret.ThirdPlaceGameRowNum = startRowNum;
			updateRequests.AddRange(CreateHeaderAndTeamRows(info, true, ref startRowIndex, ref startRowNum, "CHAMPIONSHIP: 1st-place pool winner v 2nd-place pool winner"));
			ret.ChampionshipGameRowNum = startRowNum;

			ret.UpdateValuesRequests.AddRange(updateRequests);
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
			updateRequests.Add(CreateChampionshipGameRequests(startRowIndex, homeFormula, awayFormula));
			return updateRequests;
		}

		private string GetTeamFormula(int rank, PoolPlayInfo inf)
		{
			PoolPlayInfo12Teams info = (PoolPlayInfo12Teams)inf;
			int poolWinnersStartRowNum = rank <= 3 ? info.PoolWinnersStartAndEndRowNums.First().Item1 : info.PoolWinnersStartAndEndRowNums.Last().Item1;
			int poolWinnersEndRowNum = rank <= 3 ? info.PoolWinnersStartAndEndRowNums.First().Item2 : info.PoolWinnersStartAndEndRowNums.Last().Item2;

			int rankForVlookup = rank <= 3 ? rank : 1; // rank 4 = top 2nd-place team
			return _formGen.GetTeamNameFromPoolWinnersByRankFormula(rankForVlookup, _config.TotalPoolPlayGames, poolWinnersStartRowNum, poolWinnersEndRowNum);
		}
	}
}
