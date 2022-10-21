using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public class PsoStandingsRequestCreatorConfig : ScoreBasedStandingsRequestCreatorConfig
	{
		public int NumberOfRounds { get; set; }
		public int GamesPerRound { get; set; }
	}

	public class OverallRankRequestCreatorConfig : StandingsRequestCreatorConfig
	{
		public IEnumerable<Tuple<int, int>> StandingsStartAndEndRowNums { get; set; } = Enumerable.Empty<Tuple<int, int>>();
	}
}
