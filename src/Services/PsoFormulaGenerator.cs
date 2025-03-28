﻿using System.Text;
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
		/// Gets the formula for determining the rank in the standings table by matching the team name in the sorted standings list
		/// </summary>
		/// <param name="startRowNum"></param>
		/// <param name="endRowNum"></param>
		/// <returns></returns>
		/// <remarks> =IFNA(MATCH(G3, AK$3:AK$12, 0), "")
		/// = find the team name from the sorted list of teams with tiebreakers applied
		/// </remarks>
		public string GetRankWithTiebreakerFormula(int startRowNum, int endRowNum)
			=> GetRankWithTiebreakerFormula(_helper.TeamNameColumnName, startRowNum, endRowNum);

		private string GetRankWithTiebreakerFormula(string teamNameColumnName, int startRowNum, int endRowNum)
		{
			string cellRange = Utilities.CreateCellRangeString(Utilities.ConvertIndexToColumnName(_helper.SortedStandingsListColumnIndex), startRowNum, endRowNum, CellRangeOptions.FixRow);
			string formula = string.Format("=IFNA(MATCH({0}, {1}, 0), \"\")",
				Utilities.CreateCellReference(teamNameColumnName, startRowNum),
				cellRange);
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
			=> GetRankWithTiebreakerFormula(_helper.GetColumnNameByHeader(ShootoutConstants.HDR_POOL_WINNERS), startRowNum, endRowNum);

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

		private string GetGamesPlayedCellForPoolWinnersFormulas(int startRowNum) 
			=> Utilities.CreateCellReference(_helper.GetColumnNameByHeader(ShootoutConstants.HDR_POOL_WINNER_GAMES_PLAYED), startRowNum);
		private string GetRankCellRangeForPoolWinnersFormulas(int startRowNum, int endRowNum) 
			=> Utilities.CreateCellRangeString(_helper.RankColumnName, startRowNum, endRowNum);

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
		/// Gets the formula to set up the conditional formatting to highlight the winners in the standings tables in a 5- or 10-team division
		/// </summary>
		/// <param name="rank"></param>
		/// <param name="standingsStartAndEndRowNums"></param>
		/// <param name="startRowNum"></param>
		/// <returns></returns>
		/// <remarks> =AND($O3=3,COUNTIF($G$3:$G$7, 4)=5,COUNTIF($G$24:$G$28, 4)=5) (this is for a 10-team division)
		/// = true when value of rank cell == <paramref name="rank"/>, and all 5 teams have played 4 games
		/// </remarks>
		public string GetConditionalFormattingForWinnerFormula5Teams(int rank, List<Tuple<int, int>> standingsStartAndEndRowNums, int startRowNum)
		{
			string rankCell = $"${_helper.GetColumnNameByHeader(Constants.HDR_RANK)}{startRowNum}";
			Func<Tuple<int, int>, string> getRange = pair => Utilities.CreateCellRangeString(_helper.GamesPlayedColumnName, pair.Item1, pair.Item2, CellRangeOptions.FixColumn | CellRangeOptions.FixRow);
			StringBuilder sb = new StringBuilder($"=AND({rankCell}={rank}");
			foreach (Tuple<int, int> startAndEnd in standingsStartAndEndRowNums)
			{
				string range = getRange(startAndEnd);
				sb.AppendLine($",COUNTIF({range}, {_helper.DivisionSheetConfig.TotalPoolPlayGames})={_helper.DivisionSheetConfig.TeamsPerPool}");
			}
			sb.Append(')');
			return sb.ToString();
		}

		private string GetCellRange(string columnName, IEnumerable<Tuple<int, int>> standingsStartAndEndRowNums)
		{
			string cellRanges = standingsStartAndEndRowNums.Select(item => Utilities.CreateCellRangeString(columnName, item.Item1, item.Item2, CellRangeOptions.FixRow))
						.Aggregate((s1, s2) => $"{s1},{s2}");
			if (standingsStartAndEndRowNums.Count() > 1)
				cellRanges = $"{{{cellRanges}}}";

			return cellRanges;
		}

		/// <summary>
		/// Gets the formula for calculating the result for games won tiebreaker; for PSO:
		/// Most number of wins (If a forfeit game exists for any reason, number of wins or goal differential will not be used to determine the winner.)
		/// </summary>
		/// <param name="startRow"></param>
		/// <param name="scoreEntryStartAndEndRowNums"></param>
		/// <returns>=IF(COUNTIF({E$3:E$20,E$24:E$41},TRUE)=0,I3,0)
		/// - if any of the forfeit checkboxes are checked, then 0, otherwise the number of games won
		///   (Wins not used as tiebreaker if any game in the pool/division (in case of 10 teams) has been forfeited)
		/// </returns>
		public string GetGamesWonTiebreakerFormula(int startRow, Tuple<int, int> scoreEntryStartAndEndRowNums)
		{
			string winsCell = Utilities.CreateCellReference(_helper.NumWinsColumnName, startRow);
			string noForfeitFormula = GetEnsureNoForfeitsFormula(scoreEntryStartAndEndRowNums);
			return $"=IF({noForfeitFormula}, {winsCell}, 0)";
		}

		private string GetEnsureNoForfeitsFormula(Tuple<int, int> scoreEntryStartAndEndRowNums)
		{
			string forfeitColumn = _helper.GetColumnNameByHeader(Constants.HDR_FORFEIT);
			string cellRanges = Utilities.CreateCellRangeString(forfeitColumn, scoreEntryStartAndEndRowNums.Item1, scoreEntryStartAndEndRowNums.Item2, CellRangeOptions.FixRow);
			return $"COUNTIF({cellRanges}, TRUE) = 0";
		}

		/// <summary>
		/// Gets the formula for calculating the total number of yellow and red cards for use as a tiebreaker
		/// </summary>
		/// <param name="rowNum"></param>
		/// <returns></returns>
		public string GetMisconductTiebreakerFormula(int rowNum)
		{
			string ycCell = Utilities.CreateCellReference(_helper.GetColumnNameByHeader(Constants.HDR_YELLOW_CARDS), rowNum);
			string rcCell = Utilities.CreateCellReference(_helper.GetColumnNameByHeader(Constants.HDR_RED_CARDS), rowNum);
			return $"={ycCell} + {rcCell}";
		}

		/// <summary>
		/// Gets the formula for calculating the number of goals against (allowed) for the home team for tiebreaker purposes (up to five)
		/// </summary>
		/// <param name="rowNum"></param>
		/// <returns>=IF(C3>5, 5, C3)
		/// = if the number of goals scored by the away team is greater than 5, then use 5, otherwise use the actual number of goals scored
		/// </returns>
		public string GetHomeGoalsAgainstTiebreakerFormula(int rowNum)
			=> GetGoalCountTiebreakerFormula(rowNum, _helper.AwayGoalsColumnName, 5);

		/// <summary>
		/// Gets the formula for calculating the number of goals against (allowed) for the away  team for tiebreaker purposes (up to five)
		/// </summary>
		/// <param name="rowNum"></param>
		/// <returns>=IF(B3>5, 5, B3)
		/// = if the number of goals scored by the home team is greater than 5, then use 5, otherwise use the actual number of goals scored
		/// </returns>
		public string GetAwayGoalsAgainstTiebreakerFormula(int rowNum)
			=> GetGoalCountTiebreakerFormula(rowNum, _helper.HomeGoalsColumnName, 5);

		/// <summary>
		/// Gets the formula for calculating the number of goals scored for the home team for tiebreaker purposes (up to three)
		/// </summary>
		/// <param name="rowNum"></param>
		/// <returns>=IF(B3>3, 3, B3)
		/// = if the number of goals scored by the home team is greater than 3, then use 3, otherwise use the actual number of goals scored
		/// </returns>
		public string GetHomeGoalsScoredTiebreakerFormula(int rowNum)
			=> GetGoalCountTiebreakerFormula(rowNum, _helper.HomeGoalsColumnName, 3);

		/// <summary>
		/// Gets the formula for calculating the number of goals scored for the away team for tiebreaker purposes (up to three)
		/// </summary>
		/// <param name="rowNum"></param>
		/// <returns>=IF(C3>3, 3, C3)
		/// = if the number of goals scored by the away team is greater than 3, then use 3, otherwise use the actual number of goals scored
		/// </returns>
		public string GetAwayGoalsScoredTiebreakerFormula(int rowNum)
			=> GetGoalCountTiebreakerFormula(rowNum, _helper.AwayGoalsColumnName, 3);

		private string GetGoalCountTiebreakerFormula(int rowNum, string columnName, int max)
		{
			string goalsCell = Utilities.CreateCellReference(columnName, rowNum);
			return $"=IF({goalsCell} > {max}, {max}, {goalsCell})";
		}

		/// <summary>
		/// Gets the formula for calcluating the number of goals against tiebreaker
		/// </summary>
		/// <param name="teamSheetCell"></param>
		/// <param name="scoreEntryStartAndEndRowNums"></param>
		/// <returns>=SUMIFS(AH$3:AH$12, A$3:A$12,"="&Shootout!A33, E$3:E$12, "<>TRUE")+SUMIFS(AI$3:AI$12, D$3:D$12,"="&Shootout!A33, E$3:E$12, "<>TRUE")
		/// = sum of away goals tiebreaker column where home team = team name + sum of tiebreaker home goals column where away team = team name,
		///   except when game was forfeited: if a team forfeits we have to enter the score as 1-0 and check the forfeit box, so that 1
		///   shouldn't count as a goal conceded.
		/// </returns>
		public string GetGoalsAgainstTiebreakerFormula(string teamSheetCell, Tuple<int, int> scoreEntryStartAndEndRowNums)
		{
			string formula = GetGoalCountFromTiebreakerSortSectionFormula(scoreEntryStartAndEndRowNums
						, _helper.GetColumnNameByHeader(Constants.HDR_TIEBREAKER_GOALS_AGAINST_AWAY)
						, _helper.GetColumnNameByHeader(Constants.HDR_TIEBREAKER_GOALS_AGAINST_HOME)
						, teamSheetCell
						, false
						, true);
			return $"={formula}";
		}

		/// <summary>
		/// Gets the formula for calcluating the goal differential tiebreaker
		/// </summary>
		/// <param name="teamSheetCell"></param>
		/// <param name="scoreEntryStartAndEndRowNums"></param>
		/// <returns>=SUMIFS(AF$3:AF$12, A$3:A$12,"="&Shootout!A33, {E$3:E$20, E$24:E$41}, "<>TRUE")+SUMIFS(AG$3:AG$12, D$3:D$12,"="&Shootout!A33, {E$3:E$20, E$24:E$41}, "<>TRUE") - Z3
		/// = if any of the forfeit checkboxes are checked, then 0, otherwise sum of home goals tiebreaker column where home team = team name + 
		///   sum of tiebreaker away goals column where away team = team name minus the number of goals allowed 
		///   (GD not used when any game in the pool/division (in the case of 10 teams) has been decided by forfeit)
		/// </returns>
		public string GetGoalDifferentialTiebreakerFormula(string teamSheetCell, Tuple<int, int> scoreEntryStartAndEndRowNums)
		{
			int startRow = scoreEntryStartAndEndRowNums.Item1;
			string goalsForFormula = GetGoalCountFromTiebreakerSortSectionFormula(scoreEntryStartAndEndRowNums
				, _helper.GetColumnNameByHeader(Constants.HDR_TIEBREAKER_GOALS_FOR_HOME)
				, _helper.GetColumnNameByHeader(Constants.HDR_TIEBREAKER_GOALS_FOR_AWAY)
				, teamSheetCell
				, true
				, false);
			string goalsAgainstCell = Utilities.CreateCellReference(_helper.GetColumnNameByHeader(Constants.HDR_TIEBREAKER_GOALS_AGAINST), startRow);
			string noForfeitsFormula = GetEnsureNoForfeitsFormula(scoreEntryStartAndEndRowNums);
			return $"=IF({noForfeitsFormula}, {goalsForFormula} - {goalsAgainstCell}, 0)";
		}

		private string GetGoalCountFromTiebreakerSortSectionFormula(Tuple<int, int> scoreEntryStartAndEndRowNums, string homeGoalsColName, string awayGoalsColName, string teamSheetCell, bool goalsFor, bool omitForfeits)
		{
			int startRow = scoreEntryStartAndEndRowNums.Item1;
			int endRow = scoreEntryStartAndEndRowNums.Item2;
			ScoreEntryColumns cols = new()
			{
				HomeGoalsColumnName = homeGoalsColName,
				HomeTeamColumnName = _helper.HomeTeamColumnName,
				AwayGoalsColumnName = awayGoalsColName,
				AwayTeamColumnName = _helper.AwayTeamColumnName,
			};
			return omitForfeits
				? GetGoalsFormulaOmittingForfeits(cols, startRow, endRow, teamSheetCell, goalsFor)
				: GetGoalsFormula(cols, startRow, endRow, teamSheetCell, goalsFor);
		}

		/// <summary>
		/// Gets the formula for determining the number of goals scored or conceded, excluding games which have been forfeited
		/// </summary>
		private string GetGoalsFormulaOmittingForfeits(ScoreEntryColumns scoreEntryColumns, int startRowNum, int endRowNum, string firstTeamCell, bool goalsFor)
		{
			string formula = GetGoalsFormula(scoreEntryColumns, startRowNum, endRowNum, firstTeamCell, goalsFor);
			
			string forfeitCellRange = Utilities.CreateCellRangeString(_helper.ForfeitColumnName, startRowNum, endRowNum, CellRangeOptions.FixRow);
			string unlessForfeitFormula = $", {forfeitCellRange},\"<>TRUE\"";

			// insert the unless forfeit formula into the goals formula
			int firstParen = formula.IndexOf(')');
			int lastParen = formula.LastIndexOf(')');
			string part1 = formula.Substring(0, firstParen);
			string part2 = formula.Substring(firstParen , lastParen - firstParen);
			string part3 = formula.Substring(lastParen);
			formula = string.Format("{0}{1}{2}{1}{3}", part1, unlessForfeitFormula, part2, part3);
			return formula;
		}
	}
}
