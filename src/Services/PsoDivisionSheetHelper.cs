using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
    /// <summary>
    /// An extension of <see cref="StandingsSheetHelper"/> specifically for the Panther Shootout tournament.
    /// </summary>
    /// <remarks>
    /// A separate instance must be created for each division because <see cref="_teamsPerPool"/> will vary!
    /// </remarks>
    public class PsoDivisionSheetHelper : StandingsSheetHelper
    {
        private readonly int _teamsPerPool;

		public DivisionSheetConfig DivisionSheetConfig { get; private set; }

        public PsoDivisionSheetHelper(DivisionSheetConfig config)
            : base(GameScoreColumns, StandingsHeaderRow)
        {
            _teamsPerPool = config.TeamsPerPool;
			DivisionSheetConfig = config;
        }

		#region Sheet table columns
		public static List<string> GameScoreColumns = new List<string>
		{
			Constants.HDR_HOME_TEAM,
			Constants.HDR_HOME_GOALS,
			Constants.HDR_AWAY_GOALS,
			Constants.HDR_AWAY_TEAM,
			Constants.HDR_FORFEIT,
		};
		// PSO winner column is moved off to the side so that it's not visible when the sheet is embedded
		public static List<string> WinnerAndPointsColumns = new List<string>
		{
			Constants.HDR_WINNING_TEAM,
			Constants.HDR_HOME_PTS,
			Constants.HDR_AWAY_PTS,
			Constants.HDR_TIEBREAKER_GOALS_FOR_HOME,
			Constants.HDR_TIEBREAKER_GOALS_FOR_AWAY,
			Constants.HDR_TIEBREAKER_GOALS_AGAINST_HOME,
			Constants.HDR_TIEBREAKER_GOALS_AGAINST_AWAY,
		};
		public static List<string> StandingsHeaderRow = new List<string>
		{
			Constants.HDR_RANK,
			Constants.HDR_TEAM_NAME,
			Constants.HDR_GAMES_PLAYED,
			Constants.HDR_NUM_WINS,
			Constants.HDR_NUM_LOSSES,
			Constants.HDR_NUM_DRAWS,
			Constants.HDR_YELLOW_CARDS,
			Constants.HDR_RED_CARDS,
			Constants.HDR_GAME_PTS,
			Constants.HDR_GOALS_FOR,
			Constants.HDR_GOALS_AGAINST,
			Constants.HDR_GOAL_DIFF,
		};
		public static List<string> MainTiebreakerColumns = new List<string>
		{
			Constants.HDR_TIEBREAKER_H2H,
			Constants.HDR_TIEBREAKER_WINS,
			Constants.HDR_TIEBREAKER_CARDS,
			Constants.HDR_TIEBREAKER_GOALS_AGAINST,
			Constants.HDR_TIEBREAKER_GOAL_DIFF,
			Constants.HDR_TIEBREAKER_KFTM_WINNER,
		};

		public int SortedStandingsListColumnIndex =>
			GameScoreColumns.Count +
			StandingsHeaderRow.Count +
			_teamsPerPool +
			MainTiebreakerColumns.Count +
			WinnerAndPointsColumns.Count +
			1;

		#endregion

		public override int GetColumnIndexByHeader(string colHeader)
        {
            int idx = base.GetColumnIndexByHeader(colHeader);
            if (idx > -1)
                return idx;

			idx = MainTiebreakerColumns.IndexOf(colHeader);
			if (idx > -1)
				return CalculateIndexForMainTiebreakerColumns(idx);

            idx = WinnerAndPointsColumns.IndexOf(colHeader);
            if (idx > -1)
                return CalculateIndexForAdditionalColumns(idx);

            return idx;
        }

		protected int CalculateIndexForMainTiebreakerColumns(int idx)
			=> HeaderRowColumns.Count + StandingsTableColumns.Count + _teamsPerPool + idx;

		protected int CalculateIndexForAdditionalColumns(int idx) 
			=> HeaderRowColumns.Count + StandingsTableColumns.Count + _teamsPerPool + MainTiebreakerColumns.Count + idx;

		public string HomeTeamPointsColumnName { get { return GetColumnNameByHeader(Constants.HDR_HOME_PTS); } }
        public string AwayTeamPointsColumnName { get { return GetColumnNameByHeader(Constants.HDR_AWAY_PTS); } }
        public string CalculatedRankColumnName { get { return GetColumnNameByHeader(Constants.HDR_CALC_RANK); } }
        public string Head2HeadTiebreakerColumnName { get { return GetColumnNameByHeader(Constants.HDR_TIEBREAKER_H2H); } }
		public string KicksFromTheMarkTiebreakerColumnName { get { return GetColumnNameByHeader(Constants.HDR_TIEBREAKER_KFTM_WINNER); } }
    }
}
