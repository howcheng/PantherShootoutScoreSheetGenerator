using GoogleSheetsHelper;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	/// <summary>
	/// Need one instance per division
	/// </summary>
	public class ShootoutSheetHelper : StandingsSheetHelper, IPsoSheetHelper
	{

		internal static string[] HeaderRowColumns3Rounds { get; } = new string[]
		{
			Constants.HDR_TEAM_NAME,
			ShootoutConstants.HDR_ROUND1,
			ShootoutConstants.HDR_ROUND2,
			ShootoutConstants.HDR_ROUND3,
			string.Empty,
			Constants.HDR_TOTAL_PTS,
			Constants.HDR_RANK,
		};

		internal static string[] HeaderRowColumns4Rounds { get; } = new string[]
		{
			Constants.HDR_TEAM_NAME,
			ShootoutConstants.HDR_ROUND1,
			ShootoutConstants.HDR_ROUND2,
			ShootoutConstants.HDR_ROUND3,
			ShootoutConstants.HDR_ROUND4,
			Constants.HDR_TOTAL_PTS,
			Constants.HDR_RANK,
		};

		public static readonly List<string> ScoreEntryColumns = new()
		{
			Constants.HDR_HOME_TEAM,
			Constants.HDR_HOME_GOALS,
			Constants.HDR_AWAY_GOALS,
			Constants.HDR_AWAY_TEAM,
		};

		public static readonly List<string> TiebreakerColumns = new()
		{
			ShootoutConstants.HDR_TIEBREAKER_TEAM_NAME,
			Constants.HDR_TIEBREAKER_GOALS_AGAINST,
			Constants.HDR_TIEBREAKER_KFTM_WINNER,
		};

		public DivisionSheetConfig DivisionSheetConfig { get; private set; }

		public ShootoutSheetHelper(DivisionSheetConfig config) 
			: base(config.NumberOfShootoutRounds == 4 ? HeaderRowColumns4Rounds : HeaderRowColumns3Rounds)
		{
			DivisionSheetConfig = config;
		}

		public int SortedStandingsListColumnIndex =>
			HeaderRowColumns.Count + 
			NumberOfSpacerColumnsBeforeScores +
			NumberOfShootoutRoundsColumns +
			NumberOfSpacerColumnsBeforeTiebreakers +
			TiebreakerColumns.Count + 1;

		private const int NumberOfSpacerColumnsBeforeScores = 1;
		private int NumberOfShootoutRoundsColumns => DivisionSheetConfig.NumberOfShootoutRounds * ScoreEntryColumns.Count;
		private int NumberOfSpacerColumnsBeforeTiebreakers => DivisionSheetConfig.NumberOfShootoutRounds == 3 ? ScoreEntryColumns.Count + 1 : 1;

		public override int GetColumnIndexByHeader(string colHeader)
		{
			int idx = base.GetColumnIndexByHeader(colHeader);
			if (idx > -1)
				return idx;

			idx = TiebreakerColumns.IndexOf(colHeader);
			if (idx > -1)
				return CalculateIndexForTiebreakerColumns(idx);

			return idx;
		}

		private int CalculateIndexForTiebreakerColumns(int idx)
			=> HeaderRowColumns.Count + NumberOfSpacerColumnsBeforeScores
				+ NumberOfShootoutRoundsColumns + NumberOfSpacerColumnsBeforeTiebreakers + idx;

		private int ScoreEntryStartColumnIndex => HeaderRowColumns.Count + NumberOfSpacerColumnsBeforeScores;
		public string Round1ScoreColumnName => GetColumnNameByHeader(ShootoutConstants.HDR_ROUND1);
		public string Round2ScoreColumnName => GetColumnNameByHeader(ShootoutConstants.HDR_ROUND2);
		public string Round3ScoreColumnName => GetColumnNameByHeader(ShootoutConstants.HDR_ROUND3);
		public string Round4ScoreColumnName => GetColumnNameByHeader(ShootoutConstants.HDR_ROUND4);

		public int GetColumnIndexForScoreEntry(int round) => ScoreEntryStartColumnIndex + ((round - 1) * 4);
		private string GetStartColumnNameForScoreEntry(int round)
		{
			int idx = GetColumnIndexForScoreEntry(round);
			return Utilities.ConvertIndexToColumnName(idx);
		}

		public int GetColumnIndexByHeader(string colHeader, int round)
		{
			int idx = GetColumnIndexByHeader(colHeader);
			if (idx > -1)
				return idx;

			idx = ScoreEntryColumns.IndexOf(colHeader);
			if (idx > -1)
				return idx + GetColumnIndexForScoreEntry(round);

			return idx;
		}
	}
}
