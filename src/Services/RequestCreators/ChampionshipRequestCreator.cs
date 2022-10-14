using GoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public abstract class ChampionshipRequestCreator : IChampionshipRequestCreator
	{
		protected readonly DivisionSheetConfig _config;
		protected readonly string _divisionName;
		protected readonly PsoDivisionSheetHelper _helper;

		protected ChampionshipRequestCreator(DivisionSheetConfig config, string divisionName, PsoDivisionSheetHelper helper)
		{
			_config = config;
			_divisionName = divisionName;
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
