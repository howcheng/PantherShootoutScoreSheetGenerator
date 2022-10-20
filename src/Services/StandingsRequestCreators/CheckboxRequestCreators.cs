using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public class StandingsTiebreakerRequestCreator : CheckboxRequestCreator
	{
		public StandingsTiebreakerRequestCreator(FormulaGenerator formGen) 
			: base(formGen, Constants.HDR_TIEBREAKER)
		{
		}
	}

	public class ForfeitRequestCreator : CheckboxRequestCreator
	{
		public ForfeitRequestCreator(FormulaGenerator formGen) 
			: base(formGen, Constants.HDR_FORFEIT)
		{
		}
	}

	public class PoolWinnersTiebreakerRequestCreator : CheckboxRequestCreator
	{
		public PoolWinnersTiebreakerRequestCreator(FormulaGenerator formGen)
			: base(formGen, ShootoutConstants.HDR_POOL_WINNER_TIEBREAKER)
		{
		}
	}
}
