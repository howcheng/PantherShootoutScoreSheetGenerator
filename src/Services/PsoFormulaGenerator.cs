using GoogleSheetsHelper;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public class PsoFormulaGenerator : FormulaGenerator
	{
		private readonly PsoDivisionSheetHelper _helper;

		public PsoFormulaGenerator(StandingsSheetHelper helper) : base(helper)
		{
			_helper = (PsoDivisionSheetHelper)helper;
		}

		/// <summary>
		/// Gets the formula for calculating the game points for the home team
		/// </summary>
		/// <param name="startRowNum"></param>
		/// <returns></returns>
		/// <remarks>=IFS(OR(ISBLANK(A3),P3=""),0,P3="H",6,P3="D",3,P3="A",0)+IFS(B3="",0,B3<=3,B3,B3>3,3)+IFS(C3="",0,E3=TRUE,0,,C3=0,1,C3>0,0)
		/// if team not entered yet, 0; if winner pending, 0; if home wins, 6 pts; if draw, 3 pts; if loss, 0 (can't use ISBLANK for winner cell because there's a formula there)
		/// + if home goals not entered yet, 0; if <=3 home goals, +home goals; if >3 home goals, +3
		/// + if away goals not entered yet, 0; if forfeit box is checked, 0; if 0 away goals, +1; anything else, 0
		/// </remarks>
		public string GetPsoGamePointsFormulaForHomeTeam(int startRowNum) => GetGamePointsFormula(startRowNum, true);

		/// <summary>
		/// Gets the formula for calculating the game points for the away team
		/// </summary>
		/// <param name="startRowNum"></param>
		/// <returns></returns>
		/// <remarks>=IFS(OR(ISBLANK(A3),P3=""),0,P3="A",6,P3="D",3,P3="H",0)+IFS(B3="",0,B3<=3,B3,B3>3,3)+IFS(C3="",0,E3=TRUE,0,C3=0,1,C3>0,0)
		/// if team not entered yet, 0; if winner pending, 0; if away wins, 6 pts; if draw, 3 pts; if loss, 0 (can't use ISBLANK for winner cell because there's a formula there)
		/// + if away goals not entered yet, 0; if <=3 away goals, +away goals; if >3 away goals, +3
		/// + if home goals not entered yet, 0; if forfeit box is checked, 0; if 0 home goals, +1; anything else, 0
		/// </remarks>
		public string GetPsoGamePointsFormulaForAwayTeam(int startRowNum) => GetGamePointsFormula(startRowNum, false);

		private string GetGamePointsFormula(int startRowNum, bool forHomeTeam)
		{
			string homeTeamCell = $"{_helper.HomeTeamColumnName}{startRowNum}";
			string winningTeamCell = $"{_helper.WinnerColumnName}{startRowNum}";
			string homeGoalsCell = $"{_helper.HomeGoalsColumnName}{startRowNum}";
			string awayGoalsCell = $"{_helper.AwayGoalsColumnName}{startRowNum}";
			string forfeitCell = $"{_helper.ForfeitColumnName}{startRowNum}";
			string gamePointsFormula = string.Format("=IFS(OR(ISBLANK({0}),{1}=\"\"),0,{1}=\"{5}\",6,{1}=\"{7}\",3,{1}=\"{6}\",0)+IFS(ISBLANK({3}),0,{3}<=3,{3},{3}>3,3)+IFS(ISBLANK({4}),0,{2}=TRUE,0,{4}=0,1,{4}>0,0)",
				homeTeamCell,
				winningTeamCell,
				forfeitCell,
				forHomeTeam ? homeGoalsCell : awayGoalsCell,
				forHomeTeam ? awayGoalsCell : homeGoalsCell,
				forHomeTeam ? Constants.HOME_TEAM_INDICATOR : Constants.AWAY_TEAM_INDICATOR,
				forHomeTeam ? Constants.AWAY_TEAM_INDICATOR : Constants.HOME_TEAM_INDICATOR,
				Constants.DRAW_INDICATOR
			);
			return gamePointsFormula;
		}

		/// <summary>
		/// Gets the formula for calculating the points for a round, using the calculated points listed in <see cref="PsoDivisionSheetHelper.HomeTeamPointsColumnName"/> and
		/// <see cref="PsoDivisionSheetHelper.AwayTeamPointsColumnName"/> columns
		/// </summary>
		/// <param name="startRowNum"></param>
		/// <param name="endRowNum"></param>
		/// <param name="teamsSheetCell"></param>
		/// <returns></returns>
		/// <remarks>SUMIFS(Y$3:Y$4,A$3:A$4,"="&Shootout!A53)+SUMIFS(Z$3:Z$4,D$3:D$4,"="&Shootout!A53)
		/// = sum of home points column where home team = team name + sum of away points column where away team = team name
		/// </remarks>
		public string GetTotalPointsFormula(int startRowNum, int endRowNum, string teamsSheetCell)
		{
			string homeTeamsCellRange = Utilities.CreateCellRangeString(_helper.HomeTeamColumnName, startRowNum, endRowNum, CellRangeOptions.FixRow);
			string awayTeamsCellRange = Utilities.CreateCellRangeString(_helper.AwayTeamColumnName, startRowNum, endRowNum, CellRangeOptions.FixRow);
			string homePointsCellRange = Utilities.CreateCellRangeString(_helper.HomeTeamPointsColumnName, startRowNum, endRowNum, CellRangeOptions.FixRow);
			string awayPointsCellRange = Utilities.CreateCellRangeString(_helper.AwayTeamPointsColumnName, startRowNum, endRowNum, CellRangeOptions.FixRow);
			string formula = string.Format("SUMIFS({1},{2},\"=\"&{0})+SUMIFS({3},{4},\"=\"&{0})",
				teamsSheetCell,
				homePointsCellRange,
				homeTeamsCellRange,
				awayPointsCellRange,
				awayTeamsCellRange);
			return formula;
		}

		/// <summary>
		/// Gets the formula for calculating the head-to-head winner for the home team
		/// </summary>
		/// <param name="startRowNum"></param>
		/// <param name="endRowNum"></param>
		/// <param name="firstTeamCell"></param>
		/// <param name="opponentCell"></param>
		/// <returns></returns>
		/// <remarks>SWITCH(ArrayFormula(VLOOKUP(Shootout!A17&Shootout!A18,{A$3:A$12&D$3:D$12, P$3:P$12},2,false)), "H", "W", "A", "L", "D", "D")
		/// = in the row where home team and away team are the given values, get the value from the winning team column and put "W" "L" or "D" as indicated
		/// </remarks>
		public string GetHeadToHeadComparisonFormulaForHomeTeam(int startRowNum, int endRowNum, string firstTeamCell, string opponentCell)
			=> GetHeadToHeadComparisonFormula(startRowNum, endRowNum, firstTeamCell, opponentCell, true);

		/// <summary>
		/// Gets the formula for calculating the head-to-head winner for the away team
		/// </summary>
		/// <param name="startRowNum"></param>
		/// <param name="endRowNum"></param>
		/// <param name="firstTeamCell"></param>
		/// <param name="opponentCell"></param>
		/// <returns></returns>
		/// <remarks>SWITCH(ArrayFormula(VLOOKUP(Shootout!A18&Shootout!A$17,{D$3:D$12&A$3:A$12, P$3:P$12},2,false)), "H", "L", "A", "W", "D", "D")
		/// = in the row where away team and home team are the given values, get the value from the winning team column and put "W" "L" or "D" as indicated
		/// </remarks>
		public string GetHeadToHeadComparisonFormulaForAwayTeam(int startRowNum, int endRowNum, string firstTeamCell, string opponentCell)
			=> GetHeadToHeadComparisonFormula(startRowNum, endRowNum, firstTeamCell, opponentCell, false);

		private string GetHeadToHeadComparisonFormula(int startRowNum, int endRowNum, string firstTeamCell, string opponentCell, bool forHomeTeam)
		{
			string homeTeamsRange = Utilities.CreateCellRangeString(_helper.HomeTeamColumnName, startRowNum, endRowNum, CellRangeOptions.FixRow);
			string awayTeamsRange = Utilities.CreateCellRangeString(_helper.AwayTeamColumnName, startRowNum, endRowNum, CellRangeOptions.FixRow);
			string winningTeamsRange = Utilities.CreateCellRangeString(_helper.WinnerColumnName, startRowNum, endRowNum, CellRangeOptions.FixRow);
			string teamNameCell = firstTeamCell;
			string opponentNameCell = opponentCell;

			// SWITCH(ArrayFormula(VLOOKUP(Shootout!A$17&Shootout!A18,{A$3:A$12&D$3:D$12, P$3:P$12},2,false)), "H", "W", "A", "L", "D", "D")
			// in the row where home team and away team are the given values, get the value from the winning team column and put "W" "L" or "D" as indicated
			// swap the home and away teams when forHomeTeam == false
			return string.Format("SWITCH(ArrayFormula(VLOOKUP({0}&{1},{{{2}&{3}, {4}}}, 2, FALSE)), \"{5}\", \"{8}\", \"{6}\", \"{9}\", \"{7}\", \"{7}\")",
				forHomeTeam ? teamNameCell : opponentNameCell,
				forHomeTeam ? opponentNameCell : teamNameCell,
				homeTeamsRange,
				awayTeamsRange,
				winningTeamsRange,
				Constants.HOME_TEAM_INDICATOR,
				Constants.AWAY_TEAM_INDICATOR,
				Constants.DRAW_INDICATOR,
				forHomeTeam ? Constants.WIN_INDICATOR : Constants.LOSS_INDICATOR,
				forHomeTeam ? Constants.LOSS_INDICATOR : Constants.WIN_INDICATOR
			);
		}

		/// <summary>
		/// Gets the formula for determining the rank in the standings table, taking into account the tiebreaker checkbox
		/// </summary>
		/// <param name="startRowNum"></param>
		/// <param name="endRowNum"></param>
		/// <returns></returns>
		/// <remarks>=IFS(OR(O3 >= 3, COUNTIF(O$3:O$6, O3) = 1), O3, NOT(P3), O3+1, P3, O3)
		/// = if (calculated rank >= 3 or no ties), then use the calculated rank
		/// otherwise if the tiebreaker checkbox is NOT checked, then add 1 to the calculated rank (that way, 1 becomes 2)
		/// otherwise use the calculated rank
		/// </remarks>
		public string GetRankWithTiebreakerFormula(int startRowNum, int endRowNum)
		{
			string cellRange = Utilities.CreateCellRangeString(_helper.CalculatedRankColumnName, startRowNum, endRowNum, CellRangeOptions.FixRow);
			string firstRankCell = $"{_helper.CalculatedRankColumnName}{startRowNum}";
			string firstTiebreakerCell = $"{_helper.TiebreakerColumnName}{startRowNum}";
			string rankWithTbFormula = string.Format("=IFS(OR({0} >= 3, COUNTIF({1}, {0}) = 1), {0}, NOT({2}), {0}+1, {2}, {0})",
				firstRankCell,
				cellRange,
				firstTiebreakerCell);
			return rankWithTbFormula;
		}

		/// <summary>
		/// Gets the formula for determining the rank in the standings table, taking into account the tiebreaker checkbox
		/// </summary>
		/// <param name="startRowNum"></param>
		/// <param name="endRowNum"></param>
		/// <returns></returns>
		/// <remarks>=IFNA(IFS(COUNTIF(AB$3:AB$5, AB3) = 1, AB3, NOT(AF3), AB3+1, AF3, AB3), "")
		/// = if no ties, then use the calculated rank
		/// otherwise if the tiebreaker checkbox is NOT checked, then add 1 to the calculated rank (that way, 1 becomes 2)
		/// otherwise use the calculated rank
		/// </remarks>
		public string GetPoolWinnersRankWithTiebreakerFormula(int startRowNum, int endRowNum)
		{
			string calcRankColName = _helper.GetColumnNameByHeader(ShootoutConstants.HDR_POOL_WINNER_CALC_RANK);
			string tiebreakerColName = _helper.GetColumnNameByHeader(ShootoutConstants.HDR_POOL_WINNER_TIEBREAKER);
			string cellRange = Utilities.CreateCellRangeString(calcRankColName, startRowNum, endRowNum, CellRangeOptions.FixRow);
			string firstRankCell = $"{calcRankColName}{startRowNum}";
			string firstTiebreakerCell = $"{tiebreakerColName}{startRowNum}";
			string rankWithTbFormula = string.Format("=IFNA(IFS(COUNTIF({1}, {0}) = 1, {0}, NOT({2}), {0}+1, {2}, {0}), \"\")",
				firstRankCell,
				cellRange,
				firstTiebreakerCell);
			return rankWithTbFormula;
		}

		/// <summary>
		/// Gets the formula for looking up the team name in the pool winners section of a 12-team division
		/// </summary>
		/// <param name="startRowNum"></param>
		/// <param name="standingsStartRowNum"></param>
		/// <param name="standingsEndRowNum"></param>
		/// <param name="rank"></param>
		/// <returns></returns>
		/// <remarks>=IF(W3=3, VLOOKUP(1,{M3:M6,E3:E6},2,FALSE), "")
		/// = if # of games played = 3, get the team name from the standings table where rank = 1 (or 2), else blank
		/// </remarks>
		public string GetPoolWinnersTeamNameFormula(int startRowNum, int standingsStartRowNum, int standingsEndRowNum, int rank)
		{
			string teamNameCellRange = Utilities.CreateCellRangeString(_helper.TeamNameColumnName, standingsStartRowNum, standingsEndRowNum);
			string gamesPlayedCell = GetGamesPlayedCellForPoolWinnersFormulas(startRowNum);
			string rankCellRange = GetRankCellRangeForPoolWinnersFormulas(standingsStartRowNum, standingsEndRowNum);
			return $"=IF({gamesPlayedCell}=3, VLOOKUP({rank},{{{rankCellRange},{teamNameCellRange}}},2,FALSE), \"\")";
		}

		/// <summary>
		/// Gets the formula for looking up the game points in the pool winners section of a 12-team division
		/// </summary>
		/// <param name="startRowNum"></param>
		/// <param name="standingsStartRowNum"></param>
		/// <param name="standingsEndRowNum"></param>
		/// <param name="rank"></param>
		/// <returns></returns>
		/// <remarks>=IF(W3=3, VLOOKUP(1,{M3:M6,L3:L6},2,FALSE), 0)
		/// = if # of games played = 3, get the value of the points column from the standings table where rank = 1 (or 2), else blank
		/// </remarks>
		public string GetPoolWinnersGamePointsFormula(int startRowNum, int standingsStartRowNum, int standingsEndRowNum, int rank)
		{
			string gamesPlayedCell = GetGamesPlayedCellForPoolWinnersFormulas(startRowNum);
			string rankCellRange = GetRankCellRangeForPoolWinnersFormulas(standingsStartRowNum, standingsEndRowNum);
			string pointsCellRange = Utilities.CreateCellRangeString(_helper.GamePointsColumnName, standingsStartRowNum, standingsEndRowNum);
			return $"=IF({gamesPlayedCell}=3, VLOOKUP({rank},{{{rankCellRange},{pointsCellRange}}},2,FALSE), \"\")";
		}

		/// <summary>
		/// Gets the formula for looking up the number of games played in the pool winners section of a 12-team division
		/// </summary>
		/// <param name="standingsStartRowNum"></param>
		/// <param name="standingsEndRowNum"></param>
		/// <returns></returns>
		/// <remarks>=VLOOKUP(1,{M3:M6,F3:F6},2,FALSE)
		/// = get the value of the games played column from the standings table where rank = 1 (or 2)
		/// </remarks>
		public string GetPoolWinnersGamesPlayedFormula(int standingsStartRowNum, int standingsEndRowNum)
		{
			string rankCellRange = GetRankCellRangeForPoolWinnersFormulas(standingsStartRowNum, standingsEndRowNum);
			string gamesPlayedCellRange = Utilities.CreateCellRangeString(_helper.GamesPlayedColumnName, standingsStartRowNum, standingsEndRowNum);
			return $"=VLOOKUP(1,{{{rankCellRange},{gamesPlayedCellRange}}},2,FALSE)";
		}

		private string GetGamesPlayedCellForPoolWinnersFormulas(int startRowNum) => $"{_helper.GetColumnNameByHeader(ShootoutConstants.HDR_POOL_WINNER_GAMES_PLAYED)}{startRowNum}";
		private string GetRankCellRangeForPoolWinnersFormulas(int startRowNum, int endRowNum) => Utilities.CreateCellRangeString(_helper.RankColumnName, startRowNum, endRowNum);
	}
}
