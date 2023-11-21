using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	/// <summary>
	/// Creates the checkbox to indicate that a pool play game was forfeited
	/// </summary>
	public class ForfeitRequestCreator : CheckboxRequestCreator
	{
		public ForfeitRequestCreator(FormulaGenerator formGen) 
			: base(formGen, Constants.HDR_FORFEIT)
		{
		}
	}

	public class PoolWinnersTiebreakerRequestCreator : CheckboxRequestCreator // this is probably not necessary anymore after implementing the tiebreaker algorithm
	{
		public PoolWinnersTiebreakerRequestCreator(FormulaGenerator formGen)
			: base(formGen, ShootoutConstants.HDR_POOL_WINNER_TIEBREAKER)
		{
		}
	}
}
