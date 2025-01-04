namespace PantherShootoutScoreSheetGenerator.Services
{
	public class PsoDivisionSheetHelper12Teams : PsoDivisionSheetHelper
	{
		public PsoDivisionSheetHelper12Teams(DivisionSheetConfig config) : base(config)
		{
		}

		public static readonly List<string> PoolWinnersHeaderRow = new List<string>
		{
			ShootoutConstants.HDR_POOL_WINNER_RANK,
			ShootoutConstants.HDR_POOL_WINNERS,
			ShootoutConstants.HDR_POOL_WINNER_PTS,
			ShootoutConstants.HDR_POOL_WINNER_GAMES_PLAYED,
			ShootoutConstants.HDR_POOL_WINNER_TIEBREAKER,
		};
		public static readonly List<string> RunnersUpHeaderRow = new List<string>
		{
			ShootoutConstants.HDR_POOL_WINNER_RANK,
			ShootoutConstants.HDR_RUNNERS_UP,
			ShootoutConstants.HDR_POOL_WINNER_PTS,
			ShootoutConstants.HDR_POOL_WINNER_GAMES_PLAYED,
			ShootoutConstants.HDR_POOL_WINNER_TIEBREAKER,
		};

		public override int GetColumnIndexByHeader(string colHeader)
		{
			int idx = base.GetColumnIndexByHeader(colHeader);
			if (idx > -1)
				return idx;

			idx = PoolWinnersHeaderRow.IndexOf(colHeader);
			if (idx == -1)
				return idx;

			// pool winners section starts in the same column as the winners and tiebreaker columns
			return CalculateIndexForAdditionalColumns(idx);
		}
	}
}
