using GoogleSheetsHelper;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public class ChampionshipRequestCreator4Teams : ChampionshipRequestCreator
	{
		private readonly PsoFormulaGenerator _formGen;

		public ChampionshipRequestCreator4Teams(DivisionSheetConfig config, FormulaGenerator formGen) 
			: base(config, (PsoDivisionSheetHelper)formGen.SheetHelper)
		{
			_formGen = (PsoFormulaGenerator)formGen;
		}

		/// <summary>
		/// In a 4-team division, the top two teams play for the championship and the bottom two teams play in the consolation
		/// </summary>
		/// <param name="poolPlayInfo"></param>
		/// <returns></returns>
		public override ChampionshipInfo CreateChampionshipRequests(PoolPlayInfo info)
		{
			ChampionshipInfo ret = new ChampionshipInfo(info);

			int startRowIndex = info.ChampionshipStartRowIndex;
			List<UpdateRequest> updateRequests = new List<UpdateRequest>();

			// finals label row
			UpdateRequest labelRequest = Utilities.CreateHeaderLabelRowRequest(_divisionName, startRowIndex, _helper.StandingsTableColumns.Count - 1, "FINALS", 4, cell => cell.SetHeaderCellFormatting());
			updateRequests.Add(labelRequest);
			startRowIndex += 1;

			// consolation: bottom two teams
			UpdateRequest subheaderRequest1 = Utilities.CreateHeaderLabelRowRequest(_divisionName, startRowIndex, _helper.HeaderRowColumns.Count - 1, "CONSOLATION: 3rd place vs 4th place", 0, cell => cell.SetSubheaderCellFormatting());
			updateRequests.Add(subheaderRequest1);
			startRowIndex += 1;

			UpdateRequest consolationRequest = CreateChampionshipTeamRow(info, 3, 4, startRowIndex);
			updateRequests.Add(consolationRequest);
			startRowIndex += 1;
			ret.ThirdPlaceGameRowNum = startRowIndex;

			// championship: top two teams
			UpdateRequest subheaderRequest2 = Utilities.CreateHeaderLabelRowRequest(_divisionName, startRowIndex, _helper.HeaderRowColumns.Count - 1, "CHAMPIONSHIP: 1st place vs 2nd place", 0, cell => cell.SetSubheaderCellFormatting());
			updateRequests.Add(subheaderRequest2);
			startRowIndex += 1;

			UpdateRequest championshipRequest = CreateChampionshipTeamRow(info, 1, 2, startRowIndex);
			updateRequests.Add(championshipRequest);
			ret.ChampionshipGameRowNum = startRowIndex + 1;

			ret.UpdateValuesRequests = updateRequests;
			return ret;
		}

		private UpdateRequest CreateChampionshipTeamRow(PoolPlayInfo info, int homeRank, int awayRank, int startRowIndex)
		{
			string homeFomula = _formGen.GetTeamNameFromStandingsTableByRankFormula(homeRank, _config.TotalPoolPlayGames, info.StandingsStartAndEndRowNums.First().Item1, info.StandingsStartAndEndRowNums.First().Item2);
			string awayFormula = _formGen.GetTeamNameFromStandingsTableByRankFormula(awayRank, _config.TotalPoolPlayGames, info.StandingsStartAndEndRowNums.Last().Item1, info.StandingsStartAndEndRowNums.Last().Item2);
			return CreateChampionshipGameRequests(startRowIndex, homeFomula, awayFormula);
		}
	}
}
