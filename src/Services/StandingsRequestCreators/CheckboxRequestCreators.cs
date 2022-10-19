using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public class TiebreakerRequestCreator : CheckboxRequestCreator
	{
		public TiebreakerRequestCreator(FormulaGenerator formGen) 
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
}
