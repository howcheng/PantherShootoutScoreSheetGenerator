using Microsoft.Extensions.DependencyInjection;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services.Tests
{
	public class PsoServicesServiceCollectionExtensionsTests : SheetGeneratorTests
	{
		[Theory]
		[InlineData(4)]
		[InlineData(5)]
		[InlineData(6)]
		[InlineData(8)]
		[InlineData(10)]
		[InlineData(12)]
		public void TestAddPsoServices(int numTeams)
		{
			// test that the correct services for the number of teams are added
			ServiceCollection services = new ServiceCollection();

			DivisionSheetConfig divisionConfig = DivisionSheetConfigFactory.GetForTeams(numTeams);
			List<Team> teams = CreateTeams(divisionConfig);
			ShootoutSheetConfig config = new() { TeamNameCellWidth = 100 };

			services.AddPantherShootoutServices(teams, config);
			IServiceProvider provider = services.BuildServiceProvider();

			// resolve the types that vary by number of teams
			ISortedStandingsListRequestCreator sortedStandingsCreator = provider.GetRequiredService<ISortedStandingsListRequestCreator>();
			IStandingsTableRequestCreator standingsTableCreator = provider.GetRequiredService<IStandingsTableRequestCreator>();
			DivisionSheetConfig sheetConfig = provider.GetRequiredService<DivisionSheetConfig>();
			IChampionshipRequestCreator championshipCreator = provider.GetRequiredService<IChampionshipRequestCreator>();
			StandingsSheetHelper sheetHelper = provider.GetRequiredService<StandingsSheetHelper>();
			IPoolPlayRequestCreator poolPlayCreator = provider.GetRequiredService<IPoolPlayRequestCreator>();
			IWinnerFormattingRequestsCreator winnerFormattingCreator = provider.GetRequiredService<IWinnerFormattingRequestsCreator>();

			switch (numTeams)
			{
				case 4:
					Assert.IsType<DivisionSheetConfig4Teams>(sheetConfig);
					Assert.IsType<SortedStandingsListRequestCreator>(sortedStandingsCreator);
					Assert.IsType<StandingsTableRequestCreator>(standingsTableCreator);
					Assert.IsType<ChampionshipRequestCreator4Teams>(championshipCreator);
					Assert.IsType<PsoDivisionSheetHelper>(sheetHelper);
					Assert.IsType<PoolPlayRequestCreator>(poolPlayCreator);
					Assert.IsType<WinnerFormattingRequestsCreator>(winnerFormattingCreator);
					break;
				case 5:
					Assert.IsType<DivisionSheetConfig5Teams>(sheetConfig);
					Assert.IsType<SortedStandingsListRequestCreator>(sortedStandingsCreator);
					Assert.IsType<StandingsTableRequestCreator>(standingsTableCreator);
					Assert.IsType<ChampionshipRequestCreator5Teams>(championshipCreator);
					Assert.IsType<PsoDivisionSheetHelper>(sheetHelper);
					Assert.IsType<PoolPlayRequestCreator>(poolPlayCreator);
					Assert.IsType<WinnerFormattingRequestsCreator5Teams>(winnerFormattingCreator);
					break;
				case 6:
					Assert.IsType<DivisionSheetConfig6Teams>(sheetConfig);
					Assert.IsType<SortedStandingsListRequestCreator>(sortedStandingsCreator);
					Assert.IsType<StandingsTableRequestCreator>(standingsTableCreator);
					Assert.IsType<ChampionshipRequestCreator6Teams>(championshipCreator);
					Assert.IsType<PsoDivisionSheetHelper>(sheetHelper);
					Assert.IsType<PoolPlayRequestCreator>(poolPlayCreator);
					Assert.IsType<WinnerFormattingRequestsCreator>(winnerFormattingCreator);
					break;
				case 8:
					Assert.IsType<DivisionSheetConfig8Teams>(sheetConfig);
					Assert.IsType<SortedStandingsListRequestCreator>(sortedStandingsCreator);
					Assert.IsType<StandingsTableRequestCreator>(standingsTableCreator);
					Assert.IsType<ChampionshipRequestCreator8Teams>(championshipCreator);
					Assert.IsType<PsoDivisionSheetHelper>(sheetHelper);
					Assert.IsType<PoolPlayRequestCreator>(poolPlayCreator);
					Assert.IsType<WinnerFormattingRequestsCreator>(winnerFormattingCreator);
					break;
				case 10:
					Assert.IsType<DivisionSheetConfig10Teams>(sheetConfig);
					Assert.IsType<SortedStandingsListRequestCreator10Teams>(sortedStandingsCreator);
					Assert.IsType<StandingsTableRequestCreator10Teams>(standingsTableCreator);
					Assert.IsType<ChampionshipRequestCreator5Teams>(championshipCreator);
					Assert.IsType<PsoDivisionSheetHelper>(sheetHelper);
					Assert.IsType<PoolPlayRequestCreator10Teams>(poolPlayCreator);
					Assert.IsType<WinnerFormattingRequestsCreator5Teams>(winnerFormattingCreator);
					break;
				case 12:
					Assert.IsType<DivisionSheetConfig12Teams>(sheetConfig);
					Assert.IsType<SortedStandingsListRequestCreator>(sortedStandingsCreator);
					Assert.IsType<StandingsTableRequestCreator>(standingsTableCreator);
					Assert.IsType<ChampionshipRequestCreator12Teams>(championshipCreator);
					Assert.IsType<PsoDivisionSheetHelper12Teams>(sheetHelper);
					Assert.IsType<PoolPlayRequestCreator12Teams>(poolPlayCreator);
					Assert.IsType<WinnerFormattingRequestsCreator>(winnerFormattingCreator);
					break;
			}
		}
	}

}
