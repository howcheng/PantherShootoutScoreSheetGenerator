﻿using Omu.ValueInjecter;
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

			DivisionSheetConfig configClone = new DivisionSheetConfig();
			configClone.InjectFrom(config);

			// figure out the correct types to register based on the number of teams
			Type generatorType, championshipCreatorType;
			switch (divisionTeams.Count())
			{
				case 6:
					configClone.NumberOfRounds = 3;
					configClone.GamesPerRound = 1;
					configClone.TeamsPerPool = 3;
					generatorType = typeof(SheetGenerator6Teams);
					championshipCreatorType = typeof(ChampionshipRequestCreator6Teams);
					break;
				case 8:
					configClone.NumberOfRounds = 3;
					configClone.GamesPerRound = 2;
					configClone.TeamsPerPool = 4;
					generatorType = typeof(SheetGenerator8Teams);
					championshipCreatorType = typeof(ChampionshipRequestCreator8Teams);
					break;
				case 10:
					configClone.NumberOfRounds = 5;
					configClone.GamesPerRound = 2;
					configClone.TeamsPerPool = 5;
					generatorType = typeof(SheetGenerator10Teams);
					championshipCreatorType = typeof(ChampionshipRequestCreator10Teams);
					break;
				default: // 12 teams
					configClone.NumberOfRounds = 3;
					configClone.GamesPerRound = 2;
					configClone.TeamsPerPool = 4;
					generatorType = typeof(SheetGenerator12Teams);
					championshipCreatorType = typeof(ChampionshipRequestCreator12Teams);
					break;
			}
			services.AddSingleton(configClone);
			services.AddSingleton(provider => (IDivisionSheetCreator)ActivatorUtilities.CreateInstance(provider, generatorType, divisionTeams));
			services.AddSingleton(provider => (IChampionshipRequestCreator)ActivatorUtilities.CreateInstance(provider, championshipCreatorType, divisionTeams));

			// register the correct request creators for the number of teams
			string divisionName = divisionTeams.First().DivisionName;

			PsoDivisionSheetHelper helper;
			Type poolPlayCreatorType;
			switch (divisionTeams.Count())
			{
				case 12:
					helper = new SheetHelper12Teams(configClone);
					poolPlayCreatorType = typeof(PoolPlayRequestCreator12Teams);
					break;
				default:
					helper = new PsoDivisionSheetHelper(configClone);
					poolPlayCreatorType = typeof(PoolPlayRequestCreator);
					break;
			}
			services.AddSingleton<PsoDivisionSheetHelper>(helper);
			services.AddSingleton(new FormulaGenerator(helper));
			services.AddSingleton(provider => (IPoolPlayRequestCreator)ActivatorUtilities.CreateInstance(provider, poolPlayCreatorType, divisionTeams));

			Type winnerFormatCreatorType;
			switch (divisionTeams.Count())
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
