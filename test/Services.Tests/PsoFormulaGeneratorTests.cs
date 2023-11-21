using Microsoft.Extensions.DependencyInjection;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services.Tests
{
	public class PsoFormulaGeneratorTests : SheetGeneratorTests
	{
		private const int START_ROW = 3;
		private const int END_GAMES_ROW_IN_ROUND = 4;
		private const int END_GAMES_ROW_IN_DIVISION = 14;
		private const int END_STANDINGS_ROW = 6;
		private const int END_POOL_WINNERS_ROW = 5;
		private const string TEAMS_SHEET_CELL = "Shootout!A3";

		private PsoFormulaGenerator GetFormulaGenerator(int numTeams = 4)
		{
			DivisionSheetConfig config = DivisionSheetConfigFactory.GetForTeams(numTeams);
			List<Team> teams = CreateTeams(config);
			IServiceCollection services = new ServiceCollection();
			services.AddPantherShootoutServices(teams, config);
			IServiceProvider provider = services.BuildServiceProvider();
			return (PsoFormulaGenerator)provider.GetRequiredService<FormulaGenerator>();
		}
				
		[Fact]
		public void TestGamePointsFormulaForHomeTeam()
		{
			const string expected = "=IFS(OR(ISBLANK(A3),X3=\"\"),0,X3=\"H\",6,X3=\"D\",3,X3=\"A\",0)+IFS(ISBLANK(B3),0,B3<=3,B3,B3>3,3)+IFS(ISBLANK(C3),0,E3=TRUE,0,C3=0,1,C3>0,0)";

			PsoFormulaGenerator fg = GetFormulaGenerator(4);
			string formula = fg.GetPsoGamePointsFormulaForHomeTeam(START_ROW);
			Assert.Equal(expected, formula);
		}
				
		[Fact]
		public void TestGamePointsFormulaForAwayTeam()
		{
			const string expected = "=IFS(OR(ISBLANK(A3),X3=\"\"),0,X3=\"A\",6,X3=\"D\",3,X3=\"H\",0)+IFS(ISBLANK(C3),0,C3<=3,C3,C3>3,3)+IFS(ISBLANK(B3),0,E3=TRUE,0,B3=0,1,B3>0,0)";

			PsoFormulaGenerator fg = GetFormulaGenerator(4);
			string formula = fg.GetPsoGamePointsFormulaForAwayTeam(START_ROW);
			Assert.Equal(expected, formula);
		}

		[Fact]
		public void TestTotalPointsFormula()
		{
			const string expected = "SUMIFS(Y$3:Y$4,A$3:A$4,\"=\"&Shootout!A3)+SUMIFS(Z$3:Z$4,D$3:D$4,\"=\"&Shootout!A3)";

			PsoFormulaGenerator fg = GetFormulaGenerator(4);
			string formula = fg.GetTotalPointsFormula(START_ROW, END_GAMES_ROW_IN_ROUND, TEAMS_SHEET_CELL);
			Assert.Equal(expected, formula);
		}

		[Fact]
		public void TestHeadToHeadFormulaForHomeTeam()
		{
			const string expected = "SWITCH(ArrayFormula(VLOOKUP(Shootout!A3&Shootout!A$4,{A$3:A$14&D$3:D$14, X$3:X$14}, 2, FALSE)), \"H\", \"W\", \"A\", \"L\", \"D\", \"D\")";

			PsoFormulaGenerator fg = GetFormulaGenerator(4);
			string formula = fg.GetHeadToHeadComparisonFormulaForHomeTeam(START_ROW, END_GAMES_ROW_IN_DIVISION, TEAMS_SHEET_CELL, "Shootout!A$4");
			Assert.Equal(expected, formula);
		}

		[Fact]
		public void TestHeadToHeadFormulaForAwayTeam()
		{
			const string expected = "SWITCH(ArrayFormula(VLOOKUP(Shootout!A$4&Shootout!A3,{A$3:A$14&D$3:D$14, X$3:X$14}, 2, FALSE)), \"H\", \"L\", \"A\", \"W\", \"D\", \"D\")";

			PsoFormulaGenerator fg = GetFormulaGenerator(4);
			string formula = fg.GetHeadToHeadComparisonFormulaForAwayTeam(START_ROW, END_GAMES_ROW_IN_DIVISION, TEAMS_SHEET_CELL, "Shootout!A$4");
			Assert.Equal(expected, formula);
		}

		[Fact]
		public void TestRankWithTiebreakerFormula()
		{
			const string expected = "=IFS(OR(O3 >= 3, COUNTIF(O$3:O$6, O3) = 1), O3, NOT(P3), O3+1, P3, O3)";

			PsoFormulaGenerator fg = GetFormulaGenerator(4);
			string formula = fg.GetRankWithTiebreakerFormula(END_STANDINGS_ROW);
			Assert.Equal(expected, formula);
		}

		[Fact]
		public void TestPoolWinnersRankWithTiebreakerFormula()
		{
			const string expected = "=IFNA(IFS(COUNTIF(AB$3:AB$5, AB3) = 1, AB3, NOT(AF3), AB3+1, AF3, AB3), \"\")";

			PsoFormulaGenerator fg = GetFormulaGenerator(12);
			string formula = fg.GetPoolWinnersRankWithTiebreakerFormula(START_ROW, END_POOL_WINNERS_ROW);
			Assert.Equal(expected, formula);
		}

		[Fact]
		public void TestOverallRankFormula()
		{
			const string expected = "=RANK(M3, {M$3:M$7,M$24:M$28})";

			PsoFormulaGenerator fg = GetFormulaGenerator(10);
			List<Tuple<int, int>> startAndEnd = new List<Tuple<int, int>>
			{
				new Tuple<int, int>(3, 7),
				new Tuple<int, int>(24, 28)
			};
			string formula = fg.GetOverallRankFormula(START_ROW, startAndEnd);
			Assert.Equal(expected, formula);
		}
	}
}
