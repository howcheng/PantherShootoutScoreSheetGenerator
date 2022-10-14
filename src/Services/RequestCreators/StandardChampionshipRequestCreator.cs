using GoogleSheetsHelper;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public abstract class StandardChampionshipRequestCreator : ChampionshipRequestCreator
	{
		protected StandardChampionshipRequestCreator(DivisionSheetConfig config, string divisionName, PsoDivisionSheetHelper _helper) 
			: base(config, divisionName, _helper)
		{
		}

		protected IList<UpdateRequest> CreateChampionshipHeaderAndTeamRows(PoolPlayInfo poolPlayInfo, int rank, ref int startRowIndex, ref int startRowNum, string label)
		{
			List<UpdateRequest> requests = new List<UpdateRequest>
			{
				CreateChampionshipHeaderRow(ref startRowIndex, ref startRowNum, label),
				CreateChampionshipTeamRow(poolPlayInfo, rank, ref startRowIndex, ref startRowNum),
			};
			return requests;
		}

		protected UpdateRequest CreateChampionshipHeaderRow(ref int startRowIndex, ref int startRowNum, string label)
		{
			startRowIndex += 1;
			startRowNum += 1;
			string headerCellRange = Utilities.CreateCellRangeString(_helper.HomeTeamColumnName, startRowNum, _helper.AwayTeamColumnName, startRowNum);
			UpdateRequest request = Utilities.CreateHeaderLabelRowRequest(_divisionName, startRowIndex, _helper.GetColumnIndexByHeader(Constants.HDR_AWAY_TEAM), label, 0, cell => cell.SetSubheaderCellFormatting());
			return request;
		}

		protected UpdateRequest CreateChampionshipTeamRow(PoolPlayInfo poolPlayInfo, int rank, ref int startRowIndex, ref int startRowNum)
		{
			startRowIndex += 1;
			startRowNum += 1;
			string homeFomula = GetTeamFormula(rank, poolPlayInfo.StandingsStartAndEndRowNums.First().Item1, poolPlayInfo.StandingsStartAndEndRowNums.First().Item2);
			string awayFormula = GetTeamFormula(rank, poolPlayInfo.StandingsStartAndEndRowNums.Last().Item1, poolPlayInfo.StandingsStartAndEndRowNums.Last().Item2);
			return CreateChampionshipGameRequests(startRowIndex, homeFomula, awayFormula);
		}

		protected string GetTeamFormula(int rank, int standingsStartRowNum, int standingsEndRowNum)
		{
			// =IF(COUNTIF(F3:F6,"=3")=4, VLOOKUP(1,{M3:M6,E3:E6},2,FALSE), "")
			// if not all games played yet do nothing
			// otherwise get the name of the team with rank 1
			string gamesPlayedCellRange = Utilities.CreateCellRangeString(_helper.GamesPlayedColumnName, standingsStartRowNum, standingsEndRowNum);
			string rankCellRange = Utilities.CreateCellRangeString(_helper.RankColumnName, standingsStartRowNum, standingsEndRowNum);
			string teamNameCellRange = Utilities.CreateCellRangeString(_helper.TeamNameColumnName, standingsStartRowNum, standingsEndRowNum);

			return $"=IF(COUNTIF({gamesPlayedCellRange},\"=3\")={_config.TeamsPerPool}, VLOOKUP({rank},{{{rankCellRange},{teamNameCellRange}}},2,FALSE), \"\")";
		}
	}
}
