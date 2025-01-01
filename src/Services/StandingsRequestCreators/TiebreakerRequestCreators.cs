using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public class HeadToHeadTiebreakerRequestCreator : CheckboxRequestCreator
	{
		public HeadToHeadTiebreakerRequestCreator(FormulaGenerator formGen) 
			: base(formGen, Constants.HDR_TIEBREAKER_H2H)
		{
		}
	}

	public class GamesWonTiebreakerRequestCreator : StandingsRequestCreator
	{
		public GamesWonTiebreakerRequestCreator(FormulaGenerator formGen) 
			: base(formGen, Constants.HDR_TIEBREAKER_WINS)
		{
		}

		protected override string GenerateFormula(StandingsRequestCreatorConfig cfg)
		{
			TiebreakerRequestCreatorConfig config = (TiebreakerRequestCreatorConfig)cfg;
			return ((PsoFormulaGenerator)_formulaGenerator).GetGamesWonTiebreakerFormula(config.StartGamesRowNum, config.ScoreEntryStartAndEndRowNums);
		}
	}

	public class MisconductTiebreakerRequestCreator : StandingsRequestCreator
	{
		public MisconductTiebreakerRequestCreator(FormulaGenerator formGen) 
			: base(formGen, Constants.HDR_TIEBREAKER_CARDS)
		{
		}

		protected override string GenerateFormula(StandingsRequestCreatorConfig config)
			=> ((PsoFormulaGenerator)_formulaGenerator).GetMisconductTiebreakerFormula(config.StartGamesRowNum);
	}

	public class GoalsAgainstTiebreakerRequestCreator : StandingsRequestCreator
	{
		public GoalsAgainstTiebreakerRequestCreator(FormulaGenerator formGen)
			: base(formGen, Constants.HDR_TIEBREAKER_GOALS_AGAINST)
		{
		}

		protected override string GenerateFormula(StandingsRequestCreatorConfig cfg)
		{
			TiebreakerRequestCreatorConfig config = (TiebreakerRequestCreatorConfig)cfg;
			return ((PsoFormulaGenerator)_formulaGenerator).GetGoalsAgainstTiebreakerFormula(config.FirstTeamsSheetCell, config.ScoreEntryStartAndEndRowNums);
		}
	}

	public class GoalDifferentialTiebreakerRequestCreator : StandingsRequestCreator
	{
		public GoalDifferentialTiebreakerRequestCreator(FormulaGenerator formGen) 
			: base(formGen, Constants.HDR_TIEBREAKER_GOAL_DIFF)
		{
		}

		protected override string GenerateFormula(StandingsRequestCreatorConfig cfg)
		{
			TiebreakerRequestCreatorConfig config = (TiebreakerRequestCreatorConfig)cfg;
			return ((PsoFormulaGenerator)_formulaGenerator).GetGoalDifferentialTiebreakerFormula(config.FirstTeamsSheetCell, config.ScoreEntryStartAndEndRowNums);
		}
	}

	public class KicksFromTheMarkTiebreakerRequestCreator : CheckboxRequestCreator
	{
		public KicksFromTheMarkTiebreakerRequestCreator(FormulaGenerator formGen) 
			: base(formGen, Constants.HDR_TIEBREAKER_KFTM_WINNER)
		{
		}
	}

	public class GoalsAgainstHomeTiebreakerRequestCreator : StandingsRequestCreator
	{
		public GoalsAgainstHomeTiebreakerRequestCreator(FormulaGenerator formGen) 
			: base(formGen, Constants.HDR_TIEBREAKER_GOALS_AGAINST_HOME)
		{
		}

		protected override string GenerateFormula(StandingsRequestCreatorConfig config) 
			=> ((PsoFormulaGenerator)_formulaGenerator).GetHomeGoalsAgainstTiebreakerFormula(config.StartGamesRowNum);
	}

	public class GoalsAgainstAwayTiebreakerRequestCreator : StandingsRequestCreator
	{
		public GoalsAgainstAwayTiebreakerRequestCreator(FormulaGenerator formGen) 
			: base(formGen, Constants.HDR_TIEBREAKER_GOALS_AGAINST_AWAY)
		{
		}

		protected override string GenerateFormula(StandingsRequestCreatorConfig config) 
			=> ((PsoFormulaGenerator)_formulaGenerator).GetAwayGoalsAgainstTiebreakerFormula(config.StartGamesRowNum);
	}

	public class GoalsScoredHomeTiebreakerRequestCreator : StandingsRequestCreator
	{
		public GoalsScoredHomeTiebreakerRequestCreator(FormulaGenerator formGen) 
			: base(formGen, Constants.HDR_TIEBREAKER_GOALS_FOR_HOME)
		{
		}

		protected override string GenerateFormula(StandingsRequestCreatorConfig config) 
			=> ((PsoFormulaGenerator)_formulaGenerator).GetHomeGoalsScoredTiebreakerFormula(config.StartGamesRowNum);
	}

	public class GoalsScoredAwayTiebreakerRequestCreator : StandingsRequestCreator
	{
		public GoalsScoredAwayTiebreakerRequestCreator(FormulaGenerator formGen) 
			: base(formGen, Constants.HDR_TIEBREAKER_GOALS_FOR_AWAY)
		{
		}

		protected override string GenerateFormula(StandingsRequestCreatorConfig config) 
			=> ((PsoFormulaGenerator)_formulaGenerator).GetAwayGoalsScoredTiebreakerFormula(config.StartGamesRowNum);
	}
}
