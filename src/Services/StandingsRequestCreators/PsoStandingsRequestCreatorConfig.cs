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
		public IEnumerable<Tuple<int, int>> StandingsStartAndEndRowNums { get; set; } = Enumerable.Empty<Tuple<int, int>>();
		public IEnumerable<Tuple<int, int>> ScoreEntryStartAndEndRowNums { get; set; } = Enumerable.Empty<Tuple<int, int>>();
	}
}
