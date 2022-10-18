using GoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public class ChampionshipRequestCreator8Teams : StandardChampionshipRequestCreator
	{
		public ChampionshipRequestCreator8Teams(string divisionName, DivisionSheetConfig config, PsoDivisionSheetHelper helper) 
			: base(divisionName, config, helper)
		{
		}

		/// <summary>
		/// In an 8-team division, the final match is winner of pool 1 vs winner of pool 2, and consolation is 2nd-place from pool 1 vs 2nd-place from pool 2
		/// </summary>
		/// <param name="info"></param>
		/// <returns></returns>
		public override ChampionshipInfo CreateChampionshipRequests(PoolPlayInfo info)
		{
			ChampionshipInfo ret = new ChampionshipInfo(info);

			int startRowIndex = info.ChampionshipStartRowIndex;
			int startRowNum = startRowIndex + 1;
			List<UpdateRequest> updateRequests = new List<UpdateRequest>();

			string pool1 = info.Pools!.First().First().PoolName;
			string pool2 = info.Pools!.Last().First().PoolName;

			// finals label row
			UpdateRequest headerRequest = Utilities.CreateHeaderLabelRowRequest(_divisionName, startRowIndex, _helper.HeaderRowColumns.Count - 1, "FINALS", 4, cell => cell.SetHeaderCellFormatting());
			updateRequests.Add(headerRequest);
			startRowIndex += 1;

			// championship and consolation headers and teams
			updateRequests.AddRange(CreateChampionshipHeaderAndTeamRows(info, 2, startRowIndex, startRowNum, $"3RD-PLACE: 2nd from Pool {pool1} vs 2nd from Pool {pool2}"));
			startRowIndex += 2;
			startRowNum += 2;
			ret.ThirdPlaceGameRowNum = startRowNum;
			updateRequests.AddRange(CreateChampionshipHeaderAndTeamRows(info, 1, startRowIndex, startRowNum, $"CHAMPIONSHIP: 1st from Pool {pool1} vs 1st from Pool {pool2}"));
			//startRowIndex += 2;
			startRowNum += 2;
			ret.ChampionshipGameRowNum = startRowNum;

			ret.UpdateValuesRequests = updateRequests;
			return ret;
		}
	}
}
