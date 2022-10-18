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
        public List<string> WinnerAndPointsColumns { get; }
        private readonly int _teamsPerPool;

        public PsoDivisionSheetHelper(DivisionSheetConfig config)
            : base(DivisionSheetGenerator.HeaderRowColumns, DivisionSheetGenerator.StandingsHeaderRow)
        {
            WinnerAndPointsColumns = DivisionSheetGenerator.WinnerAndPointsColumns.ToList();
            _teamsPerPool = config.TeamsPerPool;
        }

        public override int GetColumnIndexByHeader(string colHeader)
        {
            int idx = base.GetColumnIndexByHeader(colHeader);
            if (idx > -1)
                return idx;

            idx = WinnerAndPointsColumns.IndexOf(colHeader);
            if (idx > -1)
                return CalculateIndexForAdditionalColumns(idx);
            return idx;
        }

        protected int CalculateIndexForAdditionalColumns(int idx)
        {
            return HeaderRowColumns.Count + StandingsTableColumns.Count + _teamsPerPool + idx;
        }

        public string HomeTeamPointsColumnName { get { return GetColumnNameByHeader(Constants.HDR_HOME_PTS); } }
        public string AwayTeamPointsColumnName { get { return GetColumnNameByHeader(Constants.HDR_AWAY_PTS); } }
        public string CalculatedRankColumnName { get { return GetColumnNameByHeader(Constants.HDR_CALC_RANK); } }
        public string TiebreakerColumnName { get { return GetColumnNameByHeader(Constants.HDR_TIEBREAKER); } }
    }
}
