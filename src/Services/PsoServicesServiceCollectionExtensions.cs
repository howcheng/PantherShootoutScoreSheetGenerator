using Omu.ValueInjecter;
using PantherShootoutScoreSheetGenerator.Services;
using StandingsGoogleSheetsHelper;

namespace Microsoft.Extensions.DependencyInjection
{
	public static class PsoServicesServiceCollectionExtensions
	{
		public static IServiceCollection AddPantherShootoutServices(this IServiceCollection services, IEnumerable<Team> divisionTeams, DivisionSheetConfig config)
		{
			services.AddSingleton<IStandingsRequestCreatorFactory, StandingsRequestCreatorFactory>();

			// request creators for standings table
			services.AddSingleton<IStandingsRequestCreator, PsoGamesPlayedRequestCreator>();
			services.AddSingleton<IStandingsRequestCreator, PsoGamesWonRequestCreator>();
			services.AddSingleton<IStandingsRequestCreator, PsoGamesLostRequestCreator>();
			services.AddSingleton<IStandingsRequestCreator, PsoGamesDrawnRequestCreator>();
			services.AddSingleton<IStandingsRequestCreator, TotalPointsRequestCreator>();
			services.AddSingleton<IStandingsRequestCreator, RankWithTiebreakerRequestCreator>();
			services.AddSingleton<IStandingsRequestCreator, PsoTeamRankRequestCreator>();
			services.AddSingleton<IStandingsRequestCreator, GoalsScoredRequestCreator>();
			services.AddSingleton<IStandingsRequestCreator, GoalsAgainstRequestCreator>();
			services.AddSingleton<IStandingsRequestCreator, GoalDifferentialRequestCreator>();
			services.AddSingleton<IStandingsRequestCreator, HeadToHeadComparisonRequestCreator>();

			// request creators for winners and points tables
			services.AddSingleton<IStandingsRequestCreator, GameWinnerRequestCreator>();
			services.AddSingleton<IStandingsRequestCreator, HomeGamePointsRequestCreator>();
			services.AddSingleton<IStandingsRequestCreator, AwayGamePointsRequestCreator>();
			services.AddSingleton<IStandingsRequestCreator, TiebreakerRequestCreator>();
			services.AddSingleton<IStandingsRequestCreator, ForfeitRequestCreator>();

			// figure out the correct types to register based on the number of teams
			Type championshipCreatorType;
			int numTeams = divisionTeams.Count();
			switch (numTeams)
			{
				case 6:
					championshipCreatorType = typeof(ChampionshipRequestCreator6Teams);
					break;
				case 8:
					championshipCreatorType = typeof(ChampionshipRequestCreator8Teams);
					break;
				case 10:
					championshipCreatorType = typeof(ChampionshipRequestCreator10Teams);
					break;
				default: // 12 teams
					championshipCreatorType = typeof(ChampionshipRequestCreator12Teams);
					break;
			}
			DivisionSheetConfig divisionConfig = DivisionSheetConfigFactory.GetForTeams(numTeams);
			divisionConfig.InjectFrom(config);
			services.AddSingleton<DivisionSheetConfig>(divisionConfig);
			services.AddSingleton<IDivisionSheetGenerator>(provider => ActivatorUtilities.CreateInstance<DivisionSheetGenerator>(provider, divisionTeams));
			services.AddSingleton(provider => (IChampionshipRequestCreator)ActivatorUtilities.CreateInstance(provider, championshipCreatorType, divisionTeams));

			// register the correct request creators for the number of teams
			string divisionName = divisionTeams.First().DivisionName;

			PsoDivisionSheetHelper helper;
			Type poolPlayCreatorType;
			switch (numTeams)
			{
				case 12:
					helper = new PsoDivisionSheetHelper12Teams(divisionConfig);
					poolPlayCreatorType = typeof(PoolPlayRequestCreator12Teams);
					break;
				default:
					helper = new PsoDivisionSheetHelper(divisionConfig);
					poolPlayCreatorType = typeof(PoolPlayRequestCreator);
					break;
			}
			services.AddSingleton<PsoDivisionSheetHelper>(helper);
			services.AddSingleton(new FormulaGenerator(helper));
			services.AddSingleton(provider => (IPoolPlayRequestCreator)ActivatorUtilities.CreateInstance(provider, poolPlayCreatorType, divisionTeams));
			services.AddSingleton<IScoreSheetHeadersRequestCreator>(provider => ActivatorUtilities.CreateInstance<ScoreSheetHeadersRequestCreator>(provider, divisionName));
			services.AddSingleton<IScoreInputsRequestCreator>(provider => ActivatorUtilities.CreateInstance<ScoreInputsRequestCreator>(provider, divisionName));
			services.AddSingleton<IStandingsTableRequestCreator, StandingsTableRequestCreator>();

			Type winnerFormatCreatorType;
			switch (numTeams)
			{
				case 10:
					winnerFormatCreatorType = typeof(WinnerFormattingRequestsCreator10Teams);
					break;
				default:
					winnerFormatCreatorType = typeof(WinnerFormattingRequestsCreator);
					break;
			}
			services.AddSingleton(provider => (IWinnerFormattingRequestsCreator)ActivatorUtilities.CreateInstance(provider, winnerFormatCreatorType, divisionName));

			return services;
		}
	}
}
