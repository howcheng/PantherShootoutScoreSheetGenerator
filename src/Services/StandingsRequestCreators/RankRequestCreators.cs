using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	/// <summary>
	/// Used to create the rank column, based on the sorted standings list that takes into account tiebreakers
	/// </summary>
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
			int endRowNum = config.StartGamesRowNum + config.RowCount - 1; // not using EndGamesRowNum here because the pool winners table isn't the same length as the standings table
			return ((PsoFormulaGenerator)_formulaGenerator).GetRankWithTiebreakerFormula(startRowNum, endRowNum);
		}
	}

	/// <summary>
	/// Used to create the rank column for pool winners in a 12-team division, baesd on the sorted standings list that takes into account tiebreakers
	/// </summary>
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
}
