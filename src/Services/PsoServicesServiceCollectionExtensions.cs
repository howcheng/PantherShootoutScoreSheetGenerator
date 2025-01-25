using PantherShootoutScoreSheetGenerator.Services;
using StandingsGoogleSheetsHelper;

namespace Microsoft.Extensions.DependencyInjection
{
	public static class PsoServicesServiceCollectionExtensions
	{
		public static IServiceCollection AddPantherShootoutServices(this IServiceCollection services, IEnumerable<Team> divisionTeams, ShootoutSheetConfig config)
		{
			int numTeams = divisionTeams.Count();

			services.AddSingleton<IStandingsRequestCreatorFactory, StandingsRequestCreatorFactory>();

			// request creators for score entry portion
			services.AddSingleton<IStandingsRequestCreator, ForfeitRequestCreator>();

			// request creators for standings table
			services.AddSingleton<IStandingsRequestCreator, StandingsRankWithTiebreakerRequestCreator>(); 
			services.AddSingleton<IStandingsRequestCreator, PsoGamesPlayedRequestCreator>();
			services.AddSingleton<IStandingsRequestCreator, PsoGamesWonRequestCreator>();
			services.AddSingleton<IStandingsRequestCreator, PsoGamesLostRequestCreator>();
			services.AddSingleton<IStandingsRequestCreator, PsoGamesDrawnRequestCreator>();
			services.AddSingleton<IStandingsRequestCreator, TotalPointsRequestCreator>();
			services.AddSingleton<IStandingsRequestCreator, GoalsScoredRequestCreator>();
			services.AddSingleton<IStandingsRequestCreator, GoalsAgainstRequestCreator>();
			services.AddSingleton<IStandingsRequestCreator, GoalDifferentialRequestCreator>();

			// request creators for tiebreaker columns
			services.AddSingleton<ITiebreakerColumnsRequestCreator, TiebreakerColumnsRequestCreator>();
			services.AddSingleton<IStandingsRequestCreator, HeadToHeadComparisonRequestCreator>();
			services.AddSingleton<IStandingsRequestCreator, HeadToHeadTiebreakerRequestCreator>();
			services.AddSingleton<IStandingsRequestCreator, GamesWonTiebreakerRequestCreator>();
			services.AddSingleton<IStandingsRequestCreator, MisconductTiebreakerRequestCreator>();
			services.AddSingleton<IStandingsRequestCreator, GoalsAgainstTiebreakerRequestCreator>();
			services.AddSingleton<IStandingsRequestCreator, GoalDifferentialTiebreakerRequestCreator>();
			services.AddSingleton<IStandingsRequestCreator, KicksFromTheMarkTiebreakerRequestCreator>();

			// request creators for winners and points tables, and helper tiebreaker columns
			services.AddSingleton<IStandingsRequestCreator, GameWinnerRequestCreator>();
			services.AddSingleton<IStandingsRequestCreator, HomeGamePointsRequestCreator>();
			services.AddSingleton<IStandingsRequestCreator, AwayGamePointsRequestCreator>();
			services.AddSingleton<IStandingsRequestCreator, PoolWinnersRankWithTiebreakerRequestCreator>();
			services.AddSingleton<IStandingsRequestCreator, PoolWinnersTiebreakerRequestCreator>();
			services.AddSingleton<IStandingsRequestCreator, GoalsAgainstHomeTiebreakerRequestCreator>();
			services.AddSingleton<IStandingsRequestCreator, GoalsAgainstAwayTiebreakerRequestCreator>();
			services.AddSingleton<IStandingsRequestCreator, GoalsScoredHomeTiebreakerRequestCreator>();
			services.AddSingleton<IStandingsRequestCreator, GoalsScoredAwayTiebreakerRequestCreator>();

			// figure out the correct types to register based on the number of teams
			Type sortedStandingsCreatorType, standingsTableCreatorType;
			switch (numTeams)
			{
				case 10:
					sortedStandingsCreatorType = typeof(SortedStandingsListRequestCreator10Teams);
					standingsTableCreatorType = typeof(StandingsTableRequestCreator10Teams);
					break;
				default:
					sortedStandingsCreatorType = typeof (SortedStandingsListRequestCreator);
					standingsTableCreatorType = typeof(StandingsTableRequestCreator);
					break;
			}
			services.AddSingleton(typeof(ISortedStandingsListRequestCreator), sortedStandingsCreatorType);
			services.AddSingleton(typeof(IStandingsTableRequestCreator), standingsTableCreatorType);

			Type championshipCreatorType;
			switch (numTeams)
			{
				case 4:
					championshipCreatorType = typeof(ChampionshipRequestCreator4Teams);
					break;
				case 6:
					championshipCreatorType = typeof(ChampionshipRequestCreator6Teams);
					break;
				case 8:
					championshipCreatorType = typeof(ChampionshipRequestCreator8Teams);
					break;
				case 5:
				case 10:
					championshipCreatorType = typeof(ChampionshipRequestCreator5Teams);
					break;
				default: // 12 teams
					championshipCreatorType = typeof(ChampionshipRequestCreator12Teams);
					break;
			}
			DivisionSheetConfig divisionConfig = DivisionSheetConfigFactory.GetForTeams(numTeams);
			divisionConfig.TeamNameCellWidth = config.TeamNameCellWidth;
			divisionConfig.DivisionName = divisionTeams.First().DivisionName;
			services.AddSingleton(divisionConfig);
			services.AddSingleton<IDivisionSheetGenerator>(provider => ActivatorUtilities.CreateInstance<DivisionSheetGenerator>(provider, divisionTeams));
			services.AddSingleton(provider => (IChampionshipRequestCreator)ActivatorUtilities.CreateInstance(provider, championshipCreatorType));

			// register the correct request creators for the number of teams
			string divisionName = divisionTeams.First().DivisionName;

			PsoDivisionSheetHelper helper;
			Type poolPlayCreatorType;
			switch (numTeams)
			{
				case 10:
					helper = new PsoDivisionSheetHelper(divisionConfig);
					poolPlayCreatorType = typeof(PoolPlayRequestCreator10Teams);
					break;
				case 12:
					helper = new PsoDivisionSheetHelper12Teams(divisionConfig);
					poolPlayCreatorType = typeof(PoolPlayRequestCreator12Teams);
					services.AddSingleton<IPoolWinnersSortedStandingsListRequestCreator, PoolWinnersSortedStandingsListRequestCreator>();
					break;
				default:
					helper = new PsoDivisionSheetHelper(divisionConfig);
					poolPlayCreatorType = typeof(PoolPlayRequestCreator);
					break;
			}
			services.AddSingleton(helper);
			services.AddSingleton<FormulaGenerator>(new PsoFormulaGenerator(helper));
			services.AddSingleton(provider => (IPoolPlayRequestCreator)ActivatorUtilities.CreateInstance(provider, poolPlayCreatorType));
			services.AddSingleton<IScoreSheetHeadersRequestCreator, ScoreSheetHeadersRequestCreator>();
			services.AddSingleton<IScoreInputsRequestCreator, ScoreInputsRequestCreator>();

			Type winnerFormatCreatorType;
			switch (numTeams)
			{
				case 5:
				case 10:
					winnerFormatCreatorType = typeof(WinnerFormattingRequestsCreator5Teams);
					break;
				default:
					winnerFormatCreatorType = typeof(WinnerFormattingRequestsCreator);
					break;
			}
			services.AddSingleton(provider => (IWinnerFormattingRequestsCreator)ActivatorUtilities.CreateInstance(provider, winnerFormatCreatorType));

			return services;
		}
	}
}
