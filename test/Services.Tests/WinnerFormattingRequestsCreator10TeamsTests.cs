using Google.Apis.Sheets.v4.Data;
using GoogleSheetsHelper;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services.Tests
{
	public class WinnerFormattingRequestsCreator10TeamsTests : SheetGeneratorTests
	{
		[Fact]
		public void CanCreateWinnerFormattingRequestsFor10Teams()
		{
			// the only thing we need to verify in this test is the conditional formatting formula because it spans both standings tables

			DivisionSheetConfig config = DivisionSheetConfigFactory.GetForTeams(8);
			List<Team> teams = CreateTeams(config);
			PsoDivisionSheetHelper10Teams helper = new PsoDivisionSheetHelper10Teams(config);
			WinnerFormattingRequestsCreator creator = new WinnerFormattingRequestsCreator10Teams(config, new PsoFormulaGenerator(helper));

			Fixture fixture = new Fixture();
			fixture.Customize<PoolPlayInfo>(c => c.FromFactory(() => new PoolPlayInfo(teams))); // https://stackoverflow.com/questions/26149618/autofixture-customizations-provide-constructor-parameter
			PoolPlayInfo poolInfo = fixture.Build<PoolPlayInfo>()
				.Without(x => x.UpdateSheetRequests)
				.Without(x => x.UpdateValuesRequests)
				.With(x => x.StandingsStartAndEndRowNums, fixture.CreateMany<Tuple<int, int>>(2).ToList())
				.Create();
			ChampionshipInfo info = new ChampionshipInfo(poolInfo);

			SheetRequests requests = creator.CreateWinnerFormattingRequests(info);

			// the only thing we need to verify in this test is the winners conditional formatting formulas

			Assert.Equal(5, requests.UpdateSheetRequests.Count);
			IEnumerable<Request> formatRequests = requests.UpdateSheetRequests.Where(x => x.AddConditionalFormatRule != null);

			Action<GridRange, Tuple<int, int>> assertRange = (rg, pair) =>
			{
				Assert.Equal(pair.Item1 - 1, rg.StartRowIndex);
				Assert.Equal(pair.Item2, rg.EndRowIndex);
				Assert.Equal(helper.GetColumnIndexByHeader(helper.StandingsTableColumns.First()), rg.StartColumnIndex);
				Assert.Equal(helper.GetColumnIndexByHeader(helper.StandingsTableColumns.Last()) + 1, rg.EndColumnIndex);
			};
			Action<AddConditionalFormatRuleRequest, int> assertFormatRequest = (rq, rank) =>
			{
				string formula = rq.Rule.BooleanRule.Condition.Values.Single().UserEnteredValue;
				Assert.Contains($"$O{info.StandingsStartAndEndRowNums.First().Item1}", formula);
				Assert.True(Colors.GetColorForRank(rank).GoogleColorEquals(rq.Rule.BooleanRule.Format.BackgroundColor));
				Assert.Collection(rq.Rule.Ranges
					, rg => assertRange(rg, info.StandingsStartAndEndRowNums.First())
					, rg => assertRange(rg, info.StandingsStartAndEndRowNums.Last())
				);
			};

			Assert.Collection(formatRequests
				, rq => assertFormatRequest(rq.AddConditionalFormatRule, 1)
				, rq => assertFormatRequest(rq.AddConditionalFormatRule, 2)
				, rq => assertFormatRequest(rq.AddConditionalFormatRule, 3)
				, rq => assertFormatRequest(rq.AddConditionalFormatRule, 4)
				);


		}
	}
}
