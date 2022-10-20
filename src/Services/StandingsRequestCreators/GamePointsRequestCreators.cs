using Google.Apis.Sheets.v4.Data;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	/// <summary>
	/// Creates a <see cref="Request"/> for building the column for home team game points
	/// </summary>
	public class HomeGamePointsRequestCreator : StandingsRequestCreator
	{
		public HomeGamePointsRequestCreator(FormulaGenerator formGen)
			: base(formGen, Constants.HDR_HOME_PTS)
		{
		}

		protected override string GenerateFormula(StandingsRequestCreatorConfig config)
			=> ((PsoFormulaGenerator)_formulaGenerator).GetPsoGamePointsFormulaForHomeTeam(config.StartGamesRowNum);
	}

	/// <summary>
	/// Creates a <see cref="Request"/> for building the column for away team game points
	/// </summary>
	public class AwayGamePointsRequestCreator : StandingsRequestCreator
	{
		public AwayGamePointsRequestCreator(FormulaGenerator formGen)
			: base(formGen, Constants.HDR_AWAY_PTS)
		{
		}

		protected override string GenerateFormula(StandingsRequestCreatorConfig config)
			=> ((PsoFormulaGenerator)_formulaGenerator).GetPsoGamePointsFormulaForAwayTeam(config.StartGamesRowNum);
	}
}
