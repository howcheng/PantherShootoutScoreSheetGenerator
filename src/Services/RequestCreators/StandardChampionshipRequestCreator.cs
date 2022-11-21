using GoogleSheetsHelper;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public abstract class StandardChampionshipRequestCreator : ChampionshipRequestCreator
	{
		private readonly PsoFormulaGenerator _formGen;

		protected StandardChampionshipRequestCreator(DivisionSheetConfig config, FormulaGenerator formGen)
			: base(config, (PsoDivisionSheetHelper)formGen.SheetHelper)
		{
			_formGen = (PsoFormulaGenerator)formGen;
		}

		protected IList<UpdateRequest> CreateChampionshipHeaderAndTeamRows(PoolPlayInfo info, int rank, int startRowIndex, string label)
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
			string homeFomula = _formGen.GetTeamNameFromStandingsTableByRankFormula(rank, _config.TotalPoolPlayGames, info.StandingsStartAndEndRowNums.First().Item1, info.StandingsStartAndEndRowNums.First().Item2);
			string awayFormula = _formGen.GetTeamNameFromStandingsTableByRankFormula(rank, _config.TotalPoolPlayGames, info.StandingsStartAndEndRowNums.Last().Item1, info.StandingsStartAndEndRowNums.Last().Item2);
			return CreateChampionshipGameRequests(startRowIndex, homeFomula, awayFormula);
		}
	}
}
