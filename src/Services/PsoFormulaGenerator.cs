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
		/// Gets the formula for determining the team in the standings table with a given rank
		/// </summary>
		/// <param name="rank"></param>
		/// <param name="gameCount"></param>
		/// <param name="standingsStartRowNum"></param>
		/// <param name="standingsEndRowNum"></param>
		/// <returns></returns>
		/// <remarks>
		/// =IF(COUNTIF(F3:F6,"=3")=4, VLOOKUP(1,{M3:M6,E3:E6},2,FALSE), "")
		/// if not all games played yet do nothing
		/// otherwise get the name of the team with rank 1
		/// </remarks>
		public string GetTeamNameFromStandingsTableByRankFormula(int rank, int gameCount, int standingsStartRowNum, int standingsEndRowNum)
			=> GetTeamNameByRankFormula(rank, gameCount, standingsStartRowNum, standingsEndRowNum, _helper.GamesPlayedColumnName, _helper.RankColumnName, _helper.TeamNameColumnName);

		/// <summary>
		/// Gets the formula for determining the team in the pool winners section with a given rank
		/// </summary>
		/// <param name="rank"></param>
		/// <param name="gameCount"></param>
		/// <param name="standingsStartRowNum"></param>
		/// <param name="standingsEndRowNum"></param>
		/// <returns></returns>
		public string GetTeamNameFromPoolWinnersByRankFormula(int rank, int gameCount, int standingsStartRowNum, int standingsEndRowNum)
			=> GetTeamNameByRankFormula(rank, gameCount, standingsStartRowNum, standingsEndRowNum, _helper.GetColumnNameByHeader(ShootoutConstants.HDR_POOL_WINNER_GAMES_PLAYED)
				, _helper.GetColumnNameByHeader(ShootoutConstants.HDR_POOL_WINNER_RANK), _helper.GetColumnNameByHeader(ShootoutConstants.HDR_POOL_WINNERS));

		private string GetTeamNameByRankFormula(int rank, int gameCount, int standingsStartRowNum, int standingsEndRowNum, string gamesPlayedColName, string rankColName, string teamNameColName)
		{
			int teamCount = standingsEndRowNum - standingsStartRowNum + 1;
			string gamesPlayedCellRange = Utilities.CreateCellRangeString(gamesPlayedColName, standingsStartRowNum, standingsEndRowNum);
			string rankCellRange = Utilities.CreateCellRangeString(rankColName, standingsStartRowNum, standingsEndRowNum);
			string teamNameCellRange = Utilities.CreateCellRangeString(teamNameColName, standingsStartRowNum, standingsEndRowNum);

			return $"=IF(COUNTIF({gamesPlayedCellRange}, {gameCount})={teamCount}, VLOOKUP({rank},{{{rankCellRange},{teamNameCellRange}}},2,FALSE), \"\")";
		}

		/// <summary>
		/// Gets the formula for calculating the game points for the home team
		/// </summary>
		/// <param name="startRowNum"></param>
		/// <returns></returns>
		/// <remarks> =IFS(OR(ISBLANK(A3),P3=""),0,P3="H",6,P3="D",3,P3="A",0)+IFS(B3="",0,B3<=3,B3,B3>3,3)+IFS(C3="",0,E3=TRUE,0,,C3=0,1,C3>0,0)
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
		/// <remarks> =IFS(OR(ISBLANK(A3),P3=""),0,P3="A",6,P3="D",3,P3="H",0)+IFS(B3="",0,B3<=3,B3,B3>3,3)+IFS(C3="",0,E3=TRUE,0,C3=0,1,C3>0,0)
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
		/// <remarks> =IFS(OR(O3 >= 3, COUNTIF(O$3:O$6, O3) = 1), O3, NOT(P3), O3+1, P3, O3)
		/// = if (calculated rank >= 3 or no ties), then use the calculated rank
		/// otherwise if the tiebreaker checkbox is NOT checked, then add 1 to the calculated rank (that way, 1 becomes 2)
		/// otherwise use the calculated rank
		/// </remarks>
		public string GetRankWithTiebreakerFormula(int startRowNum, int endRowNum)
			=> GetRankWithTiebreakerFormula(_helper.CalculatedRankColumnName, startRowNum, endRowNum);

		/// <summary>
		/// Gets the formula for determining the rank in the standings table, taking into account the tiebreaker checkbox
		/// </summary>
		/// <param name="calcRankColumnName">Column name of the calculated rank column (for 10-team divisions, this will be "overall rank")</param>
		/// <param name="startRowNum"></param>
		/// <param name="endRowNum"></param>
		/// <returns></returns>
		/// <remarks> =IFS(OR(O3 >= 3, COUNTIF(O$3:O$6, O3) = 1), O3, NOT(P3), O3+1, P3, O3)
		/// = if (calculated rank >= 3 or no ties), then use the calculated rank
		/// otherwise if the tiebreaker checkbox is NOT checked, then add 1 to the calculated rank (that way, 1 becomes 2)
		/// otherwise use the calculated rank
		/// </remarks>
		public string GetRankWithTiebreakerFormula(string calcRankColumnName, int startRowNum, int endRowNum)
		{
			string cellRange = Utilities.CreateCellRangeString(calcRankColumnName, startRowNum, endRowNum, CellRangeOptions.FixRow);
			string firstRankCell = $"{calcRankColumnName}{startRowNum}";
			string firstTiebreakerCell = $"{_helper.TiebreakerColumnName}{startRowNum}";
			string formula = string.Format("=IFS(OR({0} >= 3, COUNTIF({1}, {0}) = 1), {0}, NOT({2}), {0}+1, {2}, {0})",
				firstRankCell,
				cellRange,
				firstTiebreakerCell);
			return formula;
		}

		/// <summary>
		/// Gets the formula for determining the rank in the standings table, taking into account the tiebreaker checkbox
		/// </summary>
		/// <param name="startRowNum"></param>
		/// <param name="endRowNum"></param>
		/// <returns></returns>
		/// <remarks> =IFNA(IFS(COUNTIF(AB$3:AB$5, AB3) = 1, AB3, NOT(AF3), AB3+1, AF3, AB3), "")
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
		/// <remarks> =IF(W3=3, VLOOKUP(1,{M3:M6,E3:E6},2,FALSE), "")
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
		/// <remarks> =IF(W3=3, VLOOKUP(1,{M3:M6,L3:L6},2,FALSE), "")
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
		/// <remarks> =IF(COUNTIF(F3:F6,3)=4, VLOOKUP(1,{M3:M6,F3:F6},2,FALSE), "")
		/// = if all teams in pool have played 3 games, get the value of the games played column from the standings table where rank = 1 (or 2), else blank
		/// </remarks>
		public string GetPoolWinnersGamesPlayedFormula(int standingsStartRowNum, int standingsEndRowNum)
		{
			string rankCellRange = GetRankCellRangeForPoolWinnersFormulas(standingsStartRowNum, standingsEndRowNum);
			string gamesPlayedCellRange = Utilities.CreateCellRangeString(_helper.GamesPlayedColumnName, standingsStartRowNum, standingsEndRowNum);
			return $"=IF(COUNTIF({gamesPlayedCellRange},3)=4, VLOOKUP(1,{{{rankCellRange},{gamesPlayedCellRange}}},2,FALSE), \"\")";
		}

		private string GetGamesPlayedCellForPoolWinnersFormulas(int startRowNum) => Utilities.CreateCellReference(_helper.GetColumnNameByHeader(ShootoutConstants.HDR_POOL_WINNER_GAMES_PLAYED), startRowNum);
		private string GetRankCellRangeForPoolWinnersFormulas(int startRowNum, int endRowNum) => Utilities.CreateCellRangeString(_helper.RankColumnName, startRowNum, endRowNum);

		/// <summary>
		/// 
		/// </summary>
		/// <returns></returns>
		/// <remarks>=IF(COUNTIF(G$3:G6, "0")=4, "", RANK(M3,M$3:M$9))
		/// = if no games played yet, show blank, otherwise do the normal rank formula
		/// </remarks>
		public string GetCalculatedRankFormula(int startRowNum, int endRowNum, int numTeams)
		{
			string rankFormula = GetTeamRankFormula(_helper.GamePointsColumnName, startRowNum, startRowNum, endRowNum);
			string gamesPlayedCellRange = Utilities.CreateCellRangeString(_helper.GamesPlayedColumnName, startRowNum, startRowNum + numTeams - 1, CellRangeOptions.FixRow);
			return $"=IF(COUNTIF({gamesPlayedCellRange}, \"0\")={numTeams}, \"\", {rankFormula})";
		}

		/// <summary>
		/// Gets the formula to set up the conditional formatting to highlight the winners in the standings tables
		/// </summary>
		/// <param name="rank"></param>
		/// <param name="gameRowNum">Row number of the championship or consolation game</param>
		/// <param name="startRowNum">Start row number of the standings table</param>
		/// <returns></returns>
		/// <remarks> =OR(AND($A$30=$E3, $B$30>$C$30), AND($D$30=$E3, $B$30<$C$30)) -- game winner
		/// = true if the team is the home team in the final game and the home score is greater than the away score, or vice versa
		/// for game losers, flip the comparion operators
		/// </remarks>
		public string GetConditionalFormattingForWinnerFormula(int rank, int gameRowNum, int startRowNum)
		{
			int standingsStartRowNum = startRowNum;
			char comparison1 = (rank % 2 == 1) ? '>' : '<';
			char comparison2 = (rank % 2 == 1) ? '<' : '>';
			return string.Format("=OR(AND(${4}${0}=${8}{1}, ${5}${0}{2}${6}${0}), AND(${7}${0}=${8}{1}, ${5}${0}{3}${6}${0}))",
				gameRowNum,
				standingsStartRowNum,
				comparison1,
				comparison2,
				_helper.HomeTeamColumnName,
				_helper.HomeGoalsColumnName,
				_helper.AwayGoalsColumnName,
				_helper.AwayTeamColumnName,
				_helper.TeamNameColumnName
			);
		}

		/// <summary>
		/// Gets the formula to set up the conditional formatting to highlight the winners in the standings tables in a 10-team division
		/// </summary>
		/// <param name="rank"></param>
		/// <param name="standingsStartAndEndRowNums"></param>
		/// <param name="startRowNum"></param>
		/// <returns></returns>
		/// <remarks> =AND($M3=1,COUNTIF($N$3:$N$7,"=4")=5,COUNTIF($N$24:$N$28,"=4")=5)
		/// = true when value of rank cell == <paramref name="rank"/>, and all teams have played 5 games
		/// </remarks>
		public string GetConditionalFormattingForWinnerFormula10Teams(int rank, List<Tuple<int, int>> standingsStartAndEndRowNums, int startRowNum)
		{
			string rankCell = $"${_helper.GetColumnNameByHeader(ShootoutConstants.HDR_OVERALL_RANK)}{startRowNum}";
			Func<Tuple<int, int>, string> getRange = pair => Utilities.CreateCellRangeString(_helper.GamesPlayedColumnName, pair.Item1, pair.Item2, CellRangeOptions.FixColumn | CellRangeOptions.FixRow);
			string pool1Range = getRange(standingsStartAndEndRowNums.First());
			string pool2Range = getRange(standingsStartAndEndRowNums.Last());
			return $"=AND({rankCell}={rank},COUNTIF({pool1Range},\"=4\")=5,COUNTIF({pool2Range},\"=4\")=5)";
		}

		/// <summary>
		/// Gets the formula for calculating the overall rank between both pools in a 10-team division
		/// </summary>
		/// <returns></returns>
		/// <remarks> =RANK(L3, {L$3:L$7,L$24:L$28})</remarks>
		public string GetOverallRankFormula(int rowNum, IEnumerable<Tuple<int, int>> standingsStartAndEndRowNums)
		{
			string columnName = _helper.GetColumnNameByHeader(Constants.HDR_GAME_PTS);
			string cellRanges = standingsStartAndEndRowNums.Select(item => Utilities.CreateCellRangeString(columnName, item.Item1, item.Item2, CellRangeOptions.FixRow))
				.Aggregate((s1, s2) => $"{s1},{s2}");
			return $"=RANK({columnName}{rowNum}, {{{cellRanges}}})";
		}
	}
}
