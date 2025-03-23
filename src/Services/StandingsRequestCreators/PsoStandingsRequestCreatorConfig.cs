using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public class PsoStandingsRequestCreatorConfig : ScoreBasedStandingsRequestCreatorConfig
	{
		public int NumberOfRounds { get; set; }
		public int GamesPerRound { get; set; }
	}

	public class TiebreakerRequestCreatorConfig : ScoreBasedStandingsRequestCreatorConfig
	{
		/// <summary>
		/// The first and last row numbers of the standings section for generating tiebreaker formulas
		/// </summary>
		public IEnumerable<Tuple<int, int>> StandingsStartAndEndRowNums { get; set; } = Enumerable.Empty<Tuple<int, int>>();
		/// <summary>
		/// THe first and last row numbers of the score entry section for generating tiebreaker formulas
		/// </summary>
		public Tuple<int, int> ScoreEntryStartAndEndRowNums { get; set; } = new(0, 0);
	}

	public class ShootoutTiebreakerRequestCreatorConfig : ScoreBasedStandingsRequestCreatorConfig
	{
		/// <summary>
		/// THe first and last row numbers of the shootout standings section for generating tiebreaker formulas
		/// </summary>
		public Tuple<int, int> StandingsStartAndEndRowNums { get; set; } = new(0, 0);
		/// <summary>
		/// THe first and last row numbers of the score entry section for generating tiebreaker formulas
		/// </summary>
		public IEnumerable<Tuple<int, int>> ScoreEntryStartAndEndRowNums { get; set; } = Enumerable.Empty<Tuple<int, int>>();
	}
}
