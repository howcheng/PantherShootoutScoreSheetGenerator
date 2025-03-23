using GoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services.Tests
{
	public class ShootoutSortedStandingsListRequestCreatorTests
	{
		[Theory]
		[InlineData(8)]
		[InlineData(10)]
		public void CanCreateSortedStandingsListRequest(int numTeams)
		{
			// confirm that the cell ranges and sort column numbers are done correctly
			DivisionSheetConfig config = DivisionSheetConfigFactory.GetForTeams(numTeams);
			ShootoutSheetHelper helper = new(config);
			ShootoutScoringFormulaGenerator fg = new(helper);

			Tuple<int, int> startAndEnd = new(3, 3 + numTeams - 1);
			ShootoutSortedStandingsListRequestCreator creator = new(fg, config);
			UpdateRequest rq = creator.CreateSortedStandingsListRequest(startAndEnd);
			string formula = rq.Rows.Last().Last().FormulaValue;

			string arrayDefn = string.Format("{{A{0}:A{1},F{0}:F{1},AA{0}:AA{1},AB{0}:AB{1}}}", startAndEnd.Item1, startAndEnd.Item2);
			Assert.Contains(arrayDefn, formula);
			Assert.Contains("},2,False", formula);

			// list should be in the same place regardless of number of teams
			Assert.Equal(29, rq.ColumnStart); // index 29 = AD
		}
	}
}
