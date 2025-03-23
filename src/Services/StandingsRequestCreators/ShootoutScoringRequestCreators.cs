using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public class ShootoutScoringShortTeamNameRequestCreator : StandingsRequestCreator
	{
		public ShootoutScoringShortTeamNameRequestCreator(FormulaGenerator fg)
			: base(fg, ShootoutConstants.HDR_TIEBREAKER_TEAM_NAME)
		{
		}

		protected override string GenerateFormula(StandingsRequestCreatorConfig config)
			=> $"=LEFT({_formulaGenerator.SheetHelper.TeamNameColumnName}{config.StartGamesRowNum},2)";
	}

	public class ShootoutScoringGoalsAgainstTiebreakerRequestCreator : StandingsRequestCreator
	{
		public ShootoutScoringGoalsAgainstTiebreakerRequestCreator(FormulaGenerator fg)
			: base(fg, Constants.HDR_TIEBREAKER_GOALS_AGAINST)
		{
		}

		protected override string GenerateFormula(StandingsRequestCreatorConfig cfg)
		{
			ShootoutTiebreakerRequestCreatorConfig config = (ShootoutTiebreakerRequestCreatorConfig)cfg;
			ShootoutSheetHelper helper = (ShootoutSheetHelper)_formulaGenerator.SheetHelper;
			ShootoutScoringFormulaGenerator formulaGenerator = (ShootoutScoringFormulaGenerator)_formulaGenerator;
			string formula = formulaGenerator.GetShootoutGoalsAgainstTiebreakerFormula(config.FirstTeamsSheetCell, config.StartGamesRowNum, config.ScoreEntryStartAndEndRowNums);
			return formula;
		}
	}
}
