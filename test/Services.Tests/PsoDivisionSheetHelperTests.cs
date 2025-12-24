using Microsoft.Extensions.DependencyInjection;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services.Tests
{
	public class PsoDivisionSheetHelperTests : SheetGeneratorTests
	{
		private PsoDivisionSheetHelper GetSheetHelper(int numTeams)
		{
			DivisionSheetConfig config = DivisionSheetConfigFactory.GetForTeams(numTeams);
			IServiceCollection services = new ServiceCollection();
			services.AddPantherShootoutServices(CreateTeams(config), new ShootoutSheetConfig());
			IServiceProvider provider = services.BuildServiceProvider();

			var helper = provider.GetRequiredService<StandingsSheetHelper>();
			return (PsoDivisionSheetHelper)helper;
		}

		[Theory]
		[InlineData(6)]
		[InlineData(8)]
		[InlineData(10)]
		[InlineData(12)]
		public void CanGetColumnIndexByHeader(int numTeams)
		{
			PsoDivisionSheetHelper helper = GetSheetHelper(numTeams);
			IEnumerable<int> indices = helper.HeaderRowColumns.Select(x => helper.GetColumnIndexByHeader(x));
			Assert.All(indices, idx => Assert.NotEqual(-1, idx));

			indices = helper.StandingsTableColumns.Select(x => helper.GetColumnIndexByHeader(x));
			Assert.All(indices, idx => Assert.NotEqual(-1, idx));

			indices = PsoDivisionSheetHelper.WinnerAndPointsColumns.Select(x => helper.GetColumnIndexByHeader(x));
			Assert.All(indices, idx => Assert.NotEqual(-1, idx));

			if (numTeams == 12)
			{
				indices = PsoDivisionSheetHelper12Teams.PoolWinnersHeaderRow.Select(x => helper.GetColumnIndexByHeader(x));
				Assert.All(indices, idx => Assert.NotEqual(-1, idx));
			}	
		}

		[Fact]
		public void TenTeamSheetHelperIsMissingCalculatedRankColumn()
		{
			PsoDivisionSheetHelper helper = GetSheetHelper(10);
			Assert.DoesNotContain(Constants.HDR_CALC_RANK, helper.StandingsTableColumns);
			Assert.Equal(-1, helper.GetColumnIndexByHeader(Constants.HDR_CALC_RANK));
		}
	}
}
