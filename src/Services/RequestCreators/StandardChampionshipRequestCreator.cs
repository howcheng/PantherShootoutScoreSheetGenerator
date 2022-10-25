using GoogleSheetsHelper;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public abstract class StandardChampionshipRequestCreator : ChampionshipRequestCreator
	{
		protected StandardChampionshipRequestCreator(DivisionSheetConfig config, PsoDivisionSheetHelper helper)
			: base(config, helper)
		{
		}

		protected IList<UpdateRequest> CreateChampionshipHeaderAndTeamRows(PoolPlayInfo info, int rank, int startRowIndex, int startRowNum, string label)
		{
			List<UpdateRequest> requests = new List<UpdateRequest>
			{
				CreateChampionshipHeaderRow(startRowIndex, label),
				CreateChampionshipTeamRow(info, rank, startRowIndex + 1),
			};
			return requests;
		}

		protected UpdateRequest CreateChampionshipHeaderRow(int startRowIndex, string label)
		{
			UpdateRequest request = Utilities.CreateHeaderLabelRowRequest(_divisionName, startRowIndex, _helper.GetColumnIndexByHeader(_helper.HeaderRowColumns.Last()), label, 0, cell => cell.SetSubheaderCellFormatting());
			return request;
		}

		protected UpdateRequest CreateChampionshipTeamRow(PoolPlayInfo info, int rank, int startRowIndex)
		{
			string homeFomula = GetTeamFormula(rank, info.StandingsStartAndEndRowNums.First().Item1, info.StandingsStartAndEndRowNums.First().Item2);
			string awayFormula = GetTeamFormula(rank, info.StandingsStartAndEndRowNums.Last().Item1, info.StandingsStartAndEndRowNums.Last().Item2);
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
