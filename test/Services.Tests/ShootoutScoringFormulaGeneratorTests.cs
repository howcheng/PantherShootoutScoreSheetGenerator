namespace PantherShootoutScoreSheetGenerator.Services.Tests
{
	public class ShootoutScoringFormulaGeneratorTests : SheetGeneratorTests
	{
		private static ShootoutScoringFormulaGenerator GetFormulaGenerator()
		{
			DivisionSheetConfig config = new();
			ShootoutSheetHelper helper = new(config);

			ShootoutScoringFormulaGenerator fg = new(helper);
			return fg;
		}

		[Theory]
		[InlineData(3)]
		[InlineData(4)]
		public void TestGetShootoutGoalsAgainstTiebreakerFormula(int numRounds)
		{
			// Goals Against formula structure (for each round):
			// SUMIFS(AwayGoals, HomeTeam, "="&team) + SUMIFS(HomeGoals, AwayTeam, "="&team)
			// = away goals where home team = you (goals you conceded as home team)
			//   + home goals where away team = you (goals you conceded as away team)
			
			// Round 1: I=HomeTeam, J=HomeGoals, K=AwayGoals, L=AwayTeam
			// Round 2: M=HomeTeam, N=HomeGoals, O=AwayGoals, P=AwayTeam
			// Round 3: Q=HomeTeam, R=HomeGoals, S=AwayGoals, T=AwayTeam
			// Round 4: U=HomeTeam, V=HomeGoals, W=AwayGoals, X=AwayTeam
			
			string expected = "=SUMIFS(K$3:K$12, I$3:I$12,\"=\"&A3)+SUMIFS(J$3:J$12, L$3:L$12,\"=\"&A3)" +
				"+SUMIFS(O$3:O$12, M$3:M$12,\"=\"&A3)+SUMIFS(N$3:N$12, P$3:P$12,\"=\"&A3)" +
				"+SUMIFS(S$3:S$12, Q$3:Q$12,\"=\"&A3)+SUMIFS(R$3:R$12, T$3:T$12,\"=\"&A3)";
			if (numRounds == 4)
				expected += "+IF(COUNT(B3:D3)<3, SUMIFS(W$3:W$12, U$3:U$12,\"=\"&A3)+SUMIFS(V$3:V$12, X$3:X$12,\"=\"&A3), 0)";

			ShootoutScoringFormulaGenerator fg = GetFormulaGenerator();

			const string teamSheetCell = "A3";
			List<Tuple<int, int>> scoringStartAndEndRows = new()
							{
								new(3, 12),
								new(3, 12),
								new(3, 12),
							};
			if (numRounds == 4)
				scoringStartAndEndRows.Add(new(3, 12));
			string formula = fg.GetShootoutGoalsAgainstTiebreakerFormula(teamSheetCell, 3, scoringStartAndEndRows);
			Assert.Equal(expected, formula);
		}

		[Fact]
		public void TestGetShootoutScoreDisplayFormula()
		{
			// Test for divisions where all teams play in every round (4, 8, or 12 teams)
			string expected = string.Format("=IF(ROWS({1}$15:{2}$16)*COLUMNS({1}$15:{2}$16)=COUNT({1}$15:{2}$16), SUMIFS({1}$15:{1}$16, {0}$15:{0}$16,\"=\"&A15)+SUMIFS({2}$15:{2}$16, {3}$15:{3}$16,\"=\"&A15), \"\")", 'I', 'J', 'K', 'L');

			ShootoutScoringFormulaGenerator fg = GetFormulaGenerator();

			const string teamSheetCell = "A15";
			const int startRowNum = 15;
			const int endRowNum = 16;
			const int roundNum = 1;
			string formula = fg.GetScoreDisplayFormula(teamSheetCell, startRowNum, endRowNum, roundNum);
			Assert.Equal(expected, formula);
		}

		[Theory]
		[InlineData(1)]
		[InlineData(4)]
		public void TestGetShootoutScoreDisplayFormulaWithBye(int roundNum)
		{
			// Test for divisions where one team has a bye each round (5/6/10 teams)
			string homeTeamCol = roundNum == 1 ? "I" : "U";
			string homeScoreCol = roundNum == 1 ? "J" : "V";
			string awayScoreCol = roundNum == 1 ? "K" : "W";
			string awayTeamCol = roundNum == 1 ? "L" : "X";

			string checkAlreadyPlayed3Turns = roundNum == 4 ? ",COUNT(B3:D3)<3" : string.Empty;
			string expected = string.Format("=IF(AND(ROWS({1}$3:{2}$4)*COLUMNS({1}$3:{2}$4)=COUNT({1}$3:{2}$4),COUNTIF({0}$3:{0}$4,\"=\"&A3)+COUNTIF({3}$3:{3}$4,\"=\"&A3)>0{4}), SUMIFS({1}$3:{1}$4, {0}$3:{0}$4,\"=\"&A3)+SUMIFS({2}$3:{2}$4, {3}$3:{3}$4,\"=\"&A3), \"\")"
				, homeTeamCol, homeScoreCol, awayScoreCol, awayTeamCol, checkAlreadyPlayed3Turns);

			ShootoutScoringFormulaGenerator fg = GetFormulaGenerator();
			const string teamSheetCell = "A3";
			const int startRowNum = 3;
			const int endRowNum = 4;
			string formula = fg.GetScoreDisplayFormulaWithBye(teamSheetCell, startRowNum, endRowNum, roundNum);
			Assert.Equal(expected, formula);
		}

		[Fact]
		public void TestGetRankFormula()
		{
			const string teamSheetCell = "A3";
			const string sortedRankCol = "AD";
			const int startRowNum = 3;
			const int endRowNum = 12;
			string expected = string.Format("=MATCH({0}, {1}${2}:{1}${3}, 0)", teamSheetCell, sortedRankCol, startRowNum, endRowNum);

			ShootoutScoringFormulaGenerator fg = GetFormulaGenerator();
			string formula = fg.GetShootoutRankFormula(teamSheetCell, startRowNum, endRowNum, sortedRankCol);
			Assert.Equal(expected, formula);
		}
	}
}