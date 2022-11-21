namespace PantherShootoutScoreSheetGenerator.Services
{
	public static class ShootoutConstants
	{
		public const string SHOOTOUT_SHEET_NAME = "Shootout";

		public const string DIV_14UG = "14U Girls";
		public const string DIV_14UB = "14U Boys";
		public const string DIV_12UG = "12U Girls";
		public const string DIV_12UB = "12U Boys";
		public const string DIV_10UG = "10U Girls";
		public const string DIV_10UB = "10U Boys";

		public static string[] DivisionNames = new string[] { DIV_10UB, DIV_10UG, DIV_12UB, DIV_12UG, DIV_14UB, DIV_14UG };

		// used in 12-team divisions
		public const string HDR_POOL_WINNERS = "Pool winners";
		public const string HDR_POOL_WINNER_PTS = "Pts*";
		public const string HDR_POOL_WINNER_RANK = "Rank*";
		public const string HDR_POOL_WINNER_CALC_RANK = "C-Rank*";
		public const string HDR_POOL_WINNER_GAMES_PLAYED = "GP*";
		public const string HDR_POOL_WINNER_TIEBREAKER = "TBW*";
		public const string HDR_RUNNERS_UP = "Runners-up";

		// used in 10-team divisions
		public const string HDR_OVERALL_RANK = "O-RANK";
	}
}