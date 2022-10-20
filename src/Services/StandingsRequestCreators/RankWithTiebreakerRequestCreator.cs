﻿using StandingsGoogleSheetsHelper;

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
			int endRowNum = config.EndGamesRowNum;
			return ((PsoFormulaGenerator)_formulaGenerator).GetPoolWinnersRankWithTiebreakerFormula(startRowNum, endRowNum);
		}
	}

	public class StandingsCalculatedRankRequestCreator : RankRequestCreator
	{
		public StandingsCalculatedRankRequestCreator(FormulaGenerator formGen)
			: base(formGen, Constants.HDR_CALC_RANK)
		{
		}
	}

	public class PoolWinnersCalculatedRankRequestCreator : StandingsRequestCreator, IStandingsRequestCreator
	{
		public PoolWinnersCalculatedRankRequestCreator(FormulaGenerator formGen) 
			: base(formGen, ShootoutConstants.HDR_POOL_WINNER_CALC_RANK)
		{
		}

		protected override string GenerateFormula(StandingsRequestCreatorConfig config) // we're just wrapping the regular formula with IFNA( ... , "")
		{
			string columnName = _formulaGenerator.SheetHelper.GetColumnNameByHeader(ShootoutConstants.HDR_POOL_WINNER_PTS);
			return $"=IFNA({_formulaGenerator.GetTeamRankFormula(columnName, config.StartGamesRowNum, config.StartGamesRowNum, config.RowCount)}, \"\")";
		}
	}
}
