using StandingsGoogleSheetsHelper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PantherShootoutScoreSheetGenerator.Services
{
    public class PsoDivisionSheetHelper : StandingsSheetHelper
    {
        public List<string> WinnerAndPointsColumns { get; }
        public int TeamsPerPool { get; set; }

        public PsoDivisionSheetHelper(IEnumerable<string> headerRowColumns, IEnumerable<string> standingsTableColumns, IEnumerable<string> winnersAndPtsColumns)
            : base(headerRowColumns, standingsTableColumns)
        {
            WinnerAndPointsColumns = winnersAndPtsColumns.ToList();
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
            return HeaderRowColumns.Count + TeamsPerPool + idx;
        }

        public string HomeTeamPointsColumnName { get { return GetColumnNameByHeader(Constants.HDR_HOME_PTS); } }
        public string AwayTeamPointsColumnName { get { return GetColumnNameByHeader(Constants.HDR_AWAY_PTS); } }
        public string CalculatedRankColumnName { get { return GetColumnNameByHeader(Constants.HDR_CALC_RANK); } }
        public string TiebreakerColumnName { get { return GetColumnNameByHeader(Constants.HDR_TIEBREAKER); } }
    }
}
