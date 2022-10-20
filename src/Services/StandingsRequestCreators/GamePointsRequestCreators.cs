using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public class HomeGamePointsRequestCreator : StandingsRequestCreator
	{
		public HomeGamePointsRequestCreator(FormulaGenerator formGen)
			: base(formGen, Constants.HDR_HOME_PTS)
		{
		}

		protected override string GenerateFormula(StandingsRequestCreatorConfig config)
			=> ((PsoFormulaGenerator)_formulaGenerator).GetPsoGamePointsFormulaForHomeTeam(config.StartGamesRowNum);
	}

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
