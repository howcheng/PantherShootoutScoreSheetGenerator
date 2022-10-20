using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public abstract class RankWithTiebreakerRequestCreator : StandingsRequestCreator, IStandingsRequestCreator
	{
		private readonly string _calculatedRankColumnName;
		private readonly string _tiebreakerColumnName;

		protected RankWithTiebreakerRequestCreator(FormulaGenerator formGen, string columnHeader, string calcRankCol, string tiebreakerCol) 
			: base(formGen, columnHeader)
		{
			_calculatedRankColumnName = calcRankCol;
			_tiebreakerColumnName = tiebreakerCol;
		}

		protected override string GenerateFormula(StandingsRequestCreatorConfig config)
		{
			int startRowNum = config.StartGamesRowNum;
			int endRowNum = config.StartGamesRowNum + config.RowCount - 1;
			return ((PsoFormulaGenerator)_formulaGenerator).GetRankWithTiebreakerFormula(startRowNum, endRowNum, _calculatedRankColumnName, _tiebreakerColumnName);
		}
	}

	public class StandingsRankWithTiebreakerRequestCreator : RankWithTiebreakerRequestCreator
	{
		public StandingsRankWithTiebreakerRequestCreator(FormulaGenerator formGen)
			: base(formGen, Constants.HDR_RANK, ((PsoDivisionSheetHelper)formGen.SheetHelper).CalculatedRankColumnName, ((PsoDivisionSheetHelper)formGen.SheetHelper).TiebreakerColumnName)
		{
		}
	}

	public class PoolWinnersRankWithTiebreakerRequestCreator : RankWithTiebreakerRequestCreator
	{
		public PoolWinnersRankWithTiebreakerRequestCreator(FormulaGenerator formGen) 
			: base(formGen, ShootoutConstants.HDR_POOL_WINNER_RANK
				, formGen.SheetHelper.GetColumnNameByHeader(ShootoutConstants.HDR_POOL_WINNER_CALC_RANK)
				, formGen.SheetHelper.GetColumnNameByHeader(ShootoutConstants.HDR_POOL_WINNER_TIEBREAKER))
		{
		}

		protected override string GenerateFormula(StandingsRequestCreatorConfig config) // we're just wrapping the regular formula with IFNA( ... , "")
		{
			string formula = base.GenerateFormula(config);
			formula = $"=IFNA({formula.Substring(1)}, \"\")";
			return formula;
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
