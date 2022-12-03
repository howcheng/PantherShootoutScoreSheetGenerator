using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public class StandingsRankWithTiebreakerRequestCreator : StandingsRequestCreator
	{
		public StandingsRankWithTiebreakerRequestCreator(FormulaGenerator formGen)
			: base(formGen, Constants.HDR_RANK)
		{
		}

		protected override string GenerateFormula(StandingsRequestCreatorConfig cfg)
		{
			PsoStandingsRequestCreatorConfig config = (PsoStandingsRequestCreatorConfig)cfg;
			int startRowNum = config.StartGamesRowNum;
			int endRowNum = config.EndGamesRowNum;
			return ((PsoFormulaGenerator)_formulaGenerator).GetRankWithTiebreakerFormula(startRowNum, endRowNum);
		}
	}

	public class PoolWinnersRankWithTiebreakerRequestCreator : StandingsRequestCreator
	{
		public PoolWinnersRankWithTiebreakerRequestCreator(FormulaGenerator formGen) 
			: base(formGen, ShootoutConstants.HDR_POOL_WINNER_RANK)
		{
		}

		protected override string GenerateFormula(StandingsRequestCreatorConfig cfg)
		{
			PsoStandingsRequestCreatorConfig config = (PsoStandingsRequestCreatorConfig)cfg;
			int startRowNum = config.StartGamesRowNum;
			int endRowNum = config.StartGamesRowNum + config.RowCount - 1; // not using EndGamesRowNum here because the pool winners table isn't the same length as the standings table
			return ((PsoFormulaGenerator)_formulaGenerator).GetPoolWinnersRankWithTiebreakerFormula(startRowNum, endRowNum);
		}
	}

	public class StandingsCalculatedRankRequestCreator : RankRequestCreator
	{
		public StandingsCalculatedRankRequestCreator(FormulaGenerator formGen)
			: base(formGen, Constants.HDR_CALC_RANK)
		{
		}

		protected override string GenerateFormula(StandingsRequestCreatorConfig cfg)
		{
			PsoStandingsRequestCreatorConfig config = (PsoStandingsRequestCreatorConfig)cfg;
			return ((PsoFormulaGenerator)_formulaGenerator).GetCalculatedRankFormula(config.StartGamesRowNum, config.EndGamesRowNum, config.RowCount);
		}
	}

	public class PoolWinnersCalculatedRankRequestCreator : StandingsRequestCreator
	{
		public PoolWinnersCalculatedRankRequestCreator(FormulaGenerator formGen) 
			: base(formGen, ShootoutConstants.HDR_POOL_WINNER_CALC_RANK)
		{
		}

		protected override string GenerateFormula(StandingsRequestCreatorConfig config) // we're just wrapping the regular formula with IFNA( ... , "")
		{
			string columnName = _formulaGenerator.SheetHelper.GetColumnNameByHeader(ShootoutConstants.HDR_POOL_WINNER_PTS);
			return $"=IFNA({_formulaGenerator.GetTeamRankFormula(columnName, config.StartGamesRowNum, config.StartGamesRowNum, config.StartGamesRowNum + config.RowCount - 1)}, \"\")";
		}
	}

	public class StandingsRankWithTiebreakerRequestCreator10Teams : RankRequestCreator
	{
		public StandingsRankWithTiebreakerRequestCreator10Teams(FormulaGenerator formGen)
			: base(formGen, Constants.HDR_RANK)
		{
		}

		protected override string GenerateFormula(StandingsRequestCreatorConfig cfg)
		{
			RankRequestCreatorConfig10Teams config = (RankRequestCreatorConfig10Teams)cfg;
			return ((PsoFormulaGenerator)_formulaGenerator).GetRankWithTiebreakerFormula10Teams(config.StartGamesRowNum, config.StandingsStartAndEndRowNums);
		}

		public override bool IsApplicableToColumn(string columnHeader) => false; // execution of this guy has to be deferred until after the standings tables are created
	}

	public class OverallRankRequestCreator : RankRequestCreator//StandingsRequestCreator
	{
		public OverallRankRequestCreator(FormulaGenerator formGen) 
			: base(formGen, ShootoutConstants.HDR_OVERALL_RANK)
		{
		}

		protected override string GenerateFormula(StandingsRequestCreatorConfig cfg)
		{
			RankRequestCreatorConfig10Teams config = (RankRequestCreatorConfig10Teams)cfg;
			return ((PsoFormulaGenerator)_formulaGenerator).GetOverallRankFormula(config.StartGamesRowNum, config.StandingsStartAndEndRowNums);
		}

		public override bool IsApplicableToColumn(string columnHeader) => false; // execution of this guy has to be deferred until after the standings tables are created
	}
}
