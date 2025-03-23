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
			string expected = "=SUMIFS({1}$3:{1}$12, {0}$3:{0}$12,\"=\"&A3)+SUMIFS({2}$3:{2}$12, {3}$3:{3}$12,\"=\"&A3)" +
				"+SUMIFS({4}$3:{4}$12, {5}$3:{5}$12,\"=\"&A3)+SUMIFS({6}$3:{6}$12, {7}$3:{7}$12,\"=\"&A3)" +
				"+SUMIFS({8}$3:{8}$12, {9}$3:{9}$12,\"=\"&A3)+SUMIFS({10}$3:{10}$12, {11}$3:{11}$12,\"=\"&A3)";
			if (numRounds == 4)
				expected += "+IF(COUNT(B3:D3)<3, SUMIFS({12}$3:{12}$12, {13}$3:{13}$12,\"=\"&A3)+SUMIFS({14}$3:{14}$12, {15}$3:{15}$12,\"=\"&A3), 0)";

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
			// test for 3 rounds (4/6/8/12 teams)
			string expected = "=IF(ROWS({1}$15:{2}$16)+COLUMNS({1}$15:{2}$16)=COUNT({1}$15:{2}$16), SUMIFS({1}$15:{1}$16, {0}$15:{0}$16,\"=\"&A15)+SUMIFS({2}$15:{2}$16, {3}$15:{3}$16,\"=\"&A15), \"\")";

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
		public void TestGetShootoutScoreDisplayFormula5Teams(int roundNum)
		{
			string homeTeamCol = roundNum == 1 ? "I" : "U";
			string homeScoreCol = roundNum == 1 ? "J" : "V";
			string awayScoreCol = roundNum == 1 ? "K" : "W";
			string awayTeamCol = roundNum == 1 ? "L" : "X";

			string checkAlreadyPlayed3Turns = roundNum == 4 ? ",COUNT(B3:D3)<3" : string.Empty;
			string expected = string.Format("=IF(AND(ROWS({1}$3:{2}$4)+COLUMNS({1}$3:{2}$4)=COUNT({1}$3:{2}$4),COUNTIF({0}$3:{0}$4,\"=\"&A3)+COUNTIF({3}$3:{3}$4,\"=\"&A3)>0{4}), SUMIFS({1}$3:{1}$4, {0}$3:{0}$4,\"=\"&A3)+SUMIFS({2}$3:{2}$4, {3}$3:{3}$4,\"=\"&A3), \"\")"
				, homeTeamCol, homeScoreCol, awayScoreCol, awayTeamCol, checkAlreadyPlayed3Turns);

			ShootoutScoringFormulaGenerator fg = GetFormulaGenerator();
			const string teamSheetCell = "A3";
			const int startRowNum = 3;
			const int endRowNum = 4;
			string formula = fg.GetScoreDisplayFormula5Teams(teamSheetCell, startRowNum, endRowNum, roundNum);
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