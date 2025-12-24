using System.Text;
using GoogleSheetsHelper;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public class ShootoutScoringFormulaGenerator : FormulaGenerator
	{
		private readonly ShootoutSheetHelper _helper;

		public ShootoutScoringFormulaGenerator(StandingsSheetHelper helper)
			: base(helper)
		{
			_helper = (ShootoutSheetHelper)helper;
		}

		/// <summary>
		/// Gets the formula for calculating the shootout goals against tiebreaker
		/// </summary>
		/// <param name="teamSheetCell"></param>
		/// <param name="scoreEntryStartAndEndRowNums"></param>
		/// <returns>=SUMIFS(K$3:K$12, I$3:I$12,"="&A3)+SUMIFS(J$3:J$12, L$3:L$12,"="&A3)
		///		+SUMIFS(O$3:O$12, M$3:M$12,"="&A3)+SUMIFS(N$3:N$12, P$3:P$12,"="&A3)
		///		+SUMIFS(S$3:S$12, Q$3:Q$12,"="&A3)+SUMIFS(R$3:R$12, T$3:T$12,"="&A3)
		///		+IF(COUNT(B3:D3)<3, SUMIFS(W$3:W$12, U$3:U$12,"="&A3)+SUMIFS(V$3:V$12, X$3:X$12,"="&A3), 0)
		///	= sum of away goals column where home team = team name + sum of home goals column where away team = team name for round 1
		///	  + same for round 2
		///	  + same for round 3
		///	  + same for round 4 if the team has a shootout in round 4 and hasn't already had 3 turns (applies only to 3- and 5-team pools / 5/6/10-team divisions)
		/// </returns>
		public string GetShootoutGoalsAgainstTiebreakerFormula(string teamSheetCell, int startRowNum, IEnumerable<Tuple<int, int>> scoreEntryStartAndEndRowNums)
		{
			StringBuilder sb = new("=");
			byte count = 0;
			foreach (var pair in scoreEntryStartAndEndRowNums)
			{
				if (count++ > 0)
					sb.Append('+');

				int homeGoalsColIdx = _helper.GetColumnIndexByHeader(Constants.HDR_HOME_GOALS, count);
				int awayGoalsColIdx = _helper.GetColumnIndexByHeader(Constants.HDR_AWAY_GOALS, count);
				int homeTeamColIdx = _helper.GetColumnIndexByHeader(Constants.HDR_HOME_TEAM, count);
				int awayTeamColIdx = _helper.GetColumnIndexByHeader(Constants.HDR_AWAY_TEAM, count);

				ScoreEntryColumns cols = new()
				{
					HomeTeamColumnName = Utilities.ConvertIndexToColumnName(homeTeamColIdx),
					HomeGoalsColumnName = Utilities.ConvertIndexToColumnName(homeGoalsColIdx),
					AwayTeamColumnName = Utilities.ConvertIndexToColumnName(awayTeamColIdx),
					AwayGoalsColumnName = Utilities.ConvertIndexToColumnName(awayGoalsColIdx),
				};

				string formulaPart = GetGoalsFormula(cols, pair.Item1, pair.Item2, teamSheetCell, goalsFor: false);
				if (count == 4)
				{
					// wrap the formula with the IF statement for round 4
					string round1cell = Utilities.CreateCellReference(_helper.Round1ScoreColumnName, startRowNum);
					string round3cell = Utilities.CreateCellReference(_helper.Round3ScoreColumnName, startRowNum);
					string entryCellRange = Utilities.CreateCellRangeString(round1cell, round3cell);
					formulaPart = $"IF(COUNT({entryCellRange})<3, {formulaPart}, 0)";
				}

				sb.Append(formulaPart);
			}
			string formula = sb.ToString();
			return formula;
		}

		/// <summary>
		/// Gets the formula for displaying the shootout score for pools of 3 or 4 teams
		/// </summary>
		/// <param name="teamSheetCell"></param>
		/// <param name="startRowNum"></param>
		/// <param name="endRowNum"></param>
		/// <param name="roundNum"></param>
		/// <returns>=IF(ROWS(J$15:K$16)*COLUMNS(J$15:K$16)=COUNT(J$15:K$16), SUMIFS(J$15:J$16, I$15:I$16,"="&A15)
		///		+SUMIFS(K$15:K$16, L$15:L$16,"="&A15), "")
		///	= if the number of rows * number of columns = the number of cells with values in the range, sum the home goals where the home team = team name
		///	  + sum the away goals where the away team = team name
		///	  otherwise blank
		/// </returns>
		public string GetScoreDisplayFormula(string teamSheetCell, int startRowNum, int endRowNum, int roundNum)
		{
			ScoreEntryColumns cols = GetScoreEntryColumnsForRound(roundNum);

			string scoreEntryCellRange = GetScoreEntryCellRangeForShootout(startRowNum, endRowNum, cols);
			string allScoresEnteredFormula = GetAllScoresEnteredForRoundFormula(scoreEntryCellRange);
			string goalsFormula = GetGoalsFormula(cols, startRowNum, endRowNum, teamSheetCell, goalsFor: true);

			string formula = $"=IF({allScoresEnteredFormula}, {goalsFormula}, \"\")";
			return formula;
		}

		/// <summary>
		/// Gets the formula for displaying the shootout score for pools of 5 teams
		/// </summary>
		/// <param name="teamSheetCell"></param>
		/// <param name="startRowNum"></param>
		/// <param name="endRowNum"></param>
		/// <param name="roundNum"></param>
		/// <returns>=IF(AND(ROWS(J$3:K$4)*COLUMNS(J$3:K$4)=COUNT(J$3:K$4),COUNTIF(I$3:I$4,"="&A3)+COUNTIF(L$3:L$4,"="&A3)>0), SUMIFS(J$3:J$4, I$3:I$4,"="&A3)
		///		+SUMIFS(K$3:K$4, L$3:L$4,"="&A3), "")
		///	= if the number of rows * number of columns = the number of cells with values in the range 
		///	  AND there is at least one cell with the team name in the home or away team column (because in a 5-team pool, 1 team will have a bye),
		///	  sum the home goals where the home team = team name + sum the away goals where the away team = team name
		///	  + sum the away goals where the away team = team name
		///	  otherwise blank
		///	  
		/// But if <paramref name="roundNum"/> is 4, the formula will be:
		/// =IF(AND(ROWS(V$3:W$4)*COLUMNS(V$3:W$4)=COUNT(V$3:W$4),COUNTIF(U$3:U$4,"="&A3)+COUNTIF(X$3:X$4,"="&A3)>0,COUNT(B3:D3)<3), SUMIFS(V$3:V$4, U$3:U$4,"="&A3)
		///		+SUMIFS(W$3:W$4, X$3:X$4,"="&A3), "")
		///	-- which is mostly the same except for the COUNT(B3:D3)<3 part because one team will have already had 3 turns so their 4th turn doesn't count
		/// </returns>
		public string GetScoreDisplayFormulaWithBye(string teamSheetCell, int startRowNum, int endRowNum, int roundNum)
		{
			ScoreEntryColumns cols = GetScoreEntryColumnsForRound(roundNum);

			string scoreEntryCellRange = GetScoreEntryCellRangeForShootout(startRowNum, endRowNum, cols);
			string allScoresEnteredFormula = GetAllScoresEnteredForRoundFormula(scoreEntryCellRange);
			string goalsFormula = GetGoalsFormula(cols, startRowNum, endRowNum, teamSheetCell, goalsFor: true);

			string homeTeamCellRange = Utilities.CreateCellRangeString(cols.HomeTeamColumnName, startRowNum, endRowNum, CellRangeOptions.FixRow);
			string awayTeamCellRange = Utilities.CreateCellRangeString(cols.AwayTeamColumnName, startRowNum, endRowNum, CellRangeOptions.FixRow);
			string checkTeamPlayedFormula = $"COUNTIF({homeTeamCellRange},\"=\"&{teamSheetCell})+COUNTIF({awayTeamCellRange},\"=\"&{teamSheetCell})>0";

			string checkAlreadyHad3TurnsFormula = string.Empty;
			if (roundNum == 4)
			{
				string round1Column = _helper.Round1ScoreColumnName;
				string round3Column = _helper.Round3ScoreColumnName;
				string round1Cell = Utilities.CreateCellReference(round1Column, startRowNum);
				string round3Cell = Utilities.CreateCellReference(round3Column, startRowNum);
				string entryCellRange = Utilities.CreateCellRangeString(round1Cell, round3Cell);
				checkAlreadyHad3TurnsFormula = $",COUNT({entryCellRange})<3";
			}

			string formula = $"=IF(AND({allScoresEnteredFormula},{checkTeamPlayedFormula}{checkAlreadyHad3TurnsFormula}), {goalsFormula}, \"\")";
			return formula;
		}

		private ScoreEntryColumns GetScoreEntryColumnsForRound(int roundNum)
		{
			int homeTeamColIdx = _helper.GetColumnIndexByHeader(Constants.HDR_HOME_TEAM, roundNum);
			int homeGoalsColIdx = _helper.GetColumnIndexByHeader(Constants.HDR_HOME_GOALS, roundNum);
			int awayGoalsColIdx = _helper.GetColumnIndexByHeader(Constants.HDR_AWAY_GOALS, roundNum);
			int awayTeamColIdx = _helper.GetColumnIndexByHeader(Constants.HDR_AWAY_TEAM, roundNum);

			string homeTeamColName = Utilities.ConvertIndexToColumnName(homeTeamColIdx);
			string homeGoalsColName = Utilities.ConvertIndexToColumnName(homeGoalsColIdx);
			string awayGoalsColName = Utilities.ConvertIndexToColumnName(awayGoalsColIdx);
			string awayTeamColName = Utilities.ConvertIndexToColumnName(awayTeamColIdx);
			return new()
			{
				HomeTeamColumnName = homeTeamColName,
				HomeGoalsColumnName = homeGoalsColName,
				AwayTeamColumnName = awayTeamColName,
				AwayGoalsColumnName = awayGoalsColName,
			};
		}

		private static string GetScoreEntryCellRangeForShootout(int startRowNum, int endRowNum, ScoreEntryColumns cols) 
			=> Utilities.CreateCellRangeString(cols.HomeGoalsColumnName, startRowNum, cols.AwayGoalsColumnName, endRowNum, CellRangeOptions.FixRow);

		private string GetAllScoresEnteredForRoundFormula(string scoreEntryCellRange)
			=> string.Format("ROWS({0})*COLUMNS({0})=COUNT({0})", scoreEntryCellRange);

		/// <summary>
		/// Gets the formula for ranking the teams in the shootout standings based on the sorted standings list
		/// </summary>
		/// <returns>=MATCH(A3, AD$3:AD$12, 0)</returns>
		public string GetShootoutRankFormula(string teamSheetCell, int startRowNum, int endRowNum, string sortedRankCol)
		{
			string cellRange = Utilities.CreateCellRangeString(sortedRankCol, startRowNum, endRowNum, CellRangeOptions.FixRow);
			return string.Format("=MATCH({0}, {1}, 0)", teamSheetCell, cellRange);
		}
	}
}
