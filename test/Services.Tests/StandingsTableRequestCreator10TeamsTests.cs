using Google.Apis.Sheets.v4.Data;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services.Tests
{
	public class StandingsTableRequestCreator10TeamsTests : SheetGeneratorTests
	{
		[Fact]
		public void ReplacesRankColumnWithCustomOne()
		{
			// test that the rank column formula is the one for 10 teams

			DivisionSheetConfig config = DivisionSheetConfigFactory.GetForTeams(10);
			List<Team> teams = CreateTeams(config);
			PsoDivisionSheetHelper helper = new(config);
			PsoFormulaGenerator fg = new(helper);

			// for this test, we only want to confirm that the rank column is done correctly
			Request req = new();
			Mock<IStandingsRequestCreator> mockReqCreator = new();
			mockReqCreator.Setup(x => x.CreateRequest(It.IsAny<StandingsRequestCreatorConfig>())).Returns(req);

			// for all other columns, return the generic mock
			Mock<IStandingsRequestCreatorFactory> mockFactory = new();
			mockFactory.Setup(x => x.GetRequestCreator(It.IsAny<string>())).Returns(mockReqCreator.Object);

			// for the rank column, return the actual class
			mockFactory.Setup(x => x.GetRequestCreator(Constants.HDR_RANK)).Returns(new StandingsRankWithTiebreakerRequestCreator(fg));

			StandingsTableRequestCreator10Teams creator = new(config, helper, mockFactory.Object);
			PoolPlayInfo info = new(teams)
			{
				StandingsStartAndEndRowNums = new()
				{
					new(3, 7),
					new(24, 28),
				},
			};
			info = creator.CreateStandingsRequests(info, info.Pools!.First(), 2);

			// verify that the rank column spans both pools
			int rankColIdx = helper.GetColumnIndexByHeader(Constants.HDR_RANK); // 5 (column F)
			IEnumerable<Request> rankRequests = info.UpdateSheetRequests.Where(x => x.RepeatCell != null && x.RepeatCell.Range.StartColumnIndex == rankColIdx);
			Assert.Equal(2, rankRequests.Count());

			Assert.All(rankRequests, rq =>
			{
				string formula = rq.RepeatCell.Cell.UserEnteredValue.FormulaValue;
				Assert.Contains("AJ$3:AJ$12", formula);
			});
		}
	}
}
