using Google.Apis.Sheets.v4.Data;
using GoogleSheetsHelper;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services.Tests
{
	public class PoolPlayRequestCreator10TeamsTests : SheetGeneratorTests
	{
		[Fact]
		public void CanCreatePoolPlayRequestsFor10Teams()
		{
			DivisionSheetConfig config = DivisionSheetConfigFactory.GetForTeams(10);
			List<Team> teams = CreateTeams(config);
			PoolPlayInfo info = new PoolPlayInfo(teams);
			Fixture fixture = new Fixture();
			info.StandingsStartAndEndRowNums = fixture.CreateMany<Tuple<int, int>>(2).ToList();

			PsoDivisionSheetHelper10Teams helper = new PsoDivisionSheetHelper10Teams(config);
			CreateMocksForPoolPlayRequestCreatorTests(config);
			IStandingsRequestCreatorFactory factory = CreateStandingsRequestCreatorFactory(teams, config);
			PoolPlayRequestCreator10Teams creator = new(config, _mockHeadersCreator.Object, _mockInputsCreator.Object, _mockStandingsCreator.Object, factory
				, _mockTiebreakerColsCreator.Object, _mockSortedStandingsCreator.Object);
			info = creator.CreatePoolPlayRequests(info);

			// verify the overall rank column
			Request? overallRankRequest = info.UpdateSheetRequests.SingleOrDefault(x => x.RepeatCell != null && x.RepeatCell.Range.StartColumnIndex == helper.GetColumnIndexByHeader(ShootoutConstants.HDR_OVERALL_RANK));
			Assert.NotNull(overallRankRequest);
			Assert.All(info.StandingsStartAndEndRowNums, startAndEnd =>
			{
				string cellRange = Utilities.CreateCellRangeString(helper.GetColumnNameByHeader(ShootoutConstants.HDR_OVERALL_RANK), startAndEnd.Item1, startAndEnd.Item2, CellRangeOptions.FixRow);
				Assert.Contains(cellRange, overallRankRequest!.RepeatCell.Cell.UserEnteredValue.FormulaValue);
			});
		}
	}
}
