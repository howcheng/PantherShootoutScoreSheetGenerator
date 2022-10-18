using GoogleSheetsHelper;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public class SheetGenerator8Teams : DivisionSheetCreator
	{
		public SheetGenerator8Teams(DivisionSheetConfig config, ISheetsClient client, FormulaGenerator fg
			, IPoolPlayRequestCreator poolPlayCreator, IChampionshipRequestCreator championshipCreator, IWinnerFormattingRequestsCreator winnerFormatCreator
			, List<Team> divisionTeams)
			: base(config, client, fg, poolPlayCreator, championshipCreator, winnerFormatCreator, divisionTeams)
		{
		}
	}

	public class ChampionshipRequestCreator8Teams : StandardChampionshipRequestCreator
	{
		public ChampionshipRequestCreator8Teams(string divisionName, DivisionSheetConfig config, PsoDivisionSheetHelper helper) 
			: base(divisionName, config, helper)
		{
		}

		/// <summary>
		/// In an 8-team division, the final match is winner of pool 1 vs winner of pool 2, and consolation is 2nd-place from pool 1 vs 2nd-place from pool 2
		/// </summary>
		/// <param name="poolPlayInfo"></param>
		/// <returns></returns>
		public override ChampionshipInfo CreateChampionshipRequests(PoolPlayInfo poolPlayInfo)
		{
			ChampionshipInfo ret = new ChampionshipInfo(poolPlayInfo);

			int startRowIndex = poolPlayInfo.ChampionshipStartRowIndex;
			int startRowNum = startRowIndex + 1;
			List<UpdateRequest> updateRequests = new List<UpdateRequest>();

			string pool1 = poolPlayInfo.Pools!.First().First().PoolName;
			string pool2 = poolPlayInfo.Pools!.Last().First().PoolName;

			// finals label row
			UpdateRequest headerRequest = Utilities.CreateHeaderLabelRowRequest(_divisionName, startRowIndex, _helper.HeaderRowColumns.Count - 1, "FINALS", 4, cell => cell.SetHeaderCellFormatting());
			updateRequests.Add(headerRequest);

			// championship and consolation headers and teams
			updateRequests.AddRange(CreateChampionshipHeaderAndTeamRows(poolPlayInfo, 2, ref startRowIndex, ref startRowNum, $"3RD-PLACE: 2nd from Pool {pool1} vs 2nd from Pool {pool2}"));
			ret.ThirdPlaceGameRowNum = startRowNum;
			updateRequests.AddRange(CreateChampionshipHeaderAndTeamRows(poolPlayInfo, 1, ref startRowIndex, ref startRowNum, $"CHAMPIONSHIP: 1st from Pool {pool1} vs 1st from Pool {pool2}"));
			ret.ChampionshipGameRowNum = startRowNum;

			ret.UpdateValuesRequests = updateRequests;
			return ret;
		}
	}
}
