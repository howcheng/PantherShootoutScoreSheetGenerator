using GoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public abstract class ChampionshipRequestCreator : IChampionshipRequestCreator
	{
		protected readonly string _divisionName;
		protected readonly DivisionSheetConfig _config;
		protected readonly PsoDivisionSheetHelper _helper;

		protected ChampionshipRequestCreator(string divisionName, DivisionSheetConfig config, PsoDivisionSheetHelper helper)
		{
			_divisionName = divisionName;
			_config = config;
			_helper = helper;
		}

		public abstract ChampionshipInfo CreateChampionshipRequests(PoolPlayInfo poolPlayInfo);

		protected UpdateRequest CreateChampionshipGameRequests(int rowIndex, string homeTeamFormula, string awayTeamFormula)
		{
			UpdateRequest request = new UpdateRequest(_divisionName)
			{
				RowStart = rowIndex
			};
			request.Rows.Add(new GoogleSheetRow
			{
				new GoogleSheetCell { FormulaValue = homeTeamFormula },
				new GoogleSheetCell(string.Empty),
				new GoogleSheetCell(string.Empty),
				new GoogleSheetCell { FormulaValue = awayTeamFormula },
			});
			return request;
		}
	}
}
