using System.Reflection;
using Google.Apis.Auth.OAuth2;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using PantherShootoutScoreSheetGenerator.Services;
using GoogleOAuthCliClient;
using GoogleSheetsHelper;
using Google.Apis.Sheets.v4;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.ConsoleApp
{
	internal class Program
	{
		private static IConfigurationRoot? Configuration;
		private static GoogleCredential? GoogleCredential;

		private static async Task Main(string[] args)
		{
			string? outputPath = Path.GetDirectoryName(new Uri(Assembly.GetExecutingAssembly().Location).LocalPath);
			Configuration = new ConfigurationBuilder()
				.SetBasePath(outputPath)
				.AddJsonFile("appSettings.json")
				.Build();

			await Run();
		}

		private static async Task Run()
		{
			AppSettings settings = new AppSettings();
			Configuration!.GetSection(nameof(AppSettings)).Bind(settings);

			IServiceProvider sharedProvider = BuildSharedServices(settings);

			// do OAuth first because we need to get the GoogleCredential
			IOAuthChecker checker = sharedProvider.GetRequiredService<IOAuthChecker>();
			if (await checker.IsOAuthRequired())
				await checker.DoOAuth();

			// load the credential
			GoogleCredential = GoogleCredential.FromAccessToken(checker.AccessToken);

			// build the service provider to read the teams and create the document
			IServiceProvider provider = BuildServiceProviderForDocument(settings);

			// read data and create the team list/shootout sheet
			ITeamReaderService teamReader = provider.GetRequiredService<ITeamReaderService>();
			IDictionary<string, IEnumerable<Team>> teams = teamReader.GetTeams();

			// create the spreadsheet
			ISheetsClient sheetsClient = provider.GetRequiredService<ISheetsClient>();
			await sheetsClient.CreateSpreadsheet($"{DateTime.Today.Year} Panther Shootout scores");
			ILogger<Program> logger = provider.GetRequiredService<ILogger<Program>>();
			logger.LogInformation("Spreadsheet ID {id} created", sheetsClient.SpreadsheetId);

			// Create division configs for all divisions because we need them to create the Shootout sheet
			Dictionary<string, DivisionSheetConfig> divisionConfigs = new Dictionary<string, DivisionSheetConfig>();
			foreach (KeyValuePair<string, IEnumerable<Team>> pair in teams)
			{
				DivisionSheetConfig divisionConfig = DivisionSheetConfigFactory.GetForTeams(pair.Value.Count());
				divisionConfig.DivisionName = pair.Key;
				divisionConfigs.Add(pair.Key, divisionConfig);
			}

			IShootoutSheetService shootoutSheetService = provider.GetRequiredService<IShootoutSheetService>();
			ShootoutSheetConfig config = await shootoutSheetService.GenerateSheet(teams, divisionConfigs);

			// Update all division configs with the team name cell width from the shootout sheet
			foreach (var divConfig in config.DivisionConfigs.Values)
			{
				divConfig.TeamNameCellWidth = config.TeamNameCellWidth;
			}

			// loop thru the divisions
			foreach (KeyValuePair<string, IEnumerable<Team>> pair in teams)
			{
				// build the provider for the division and create the sheet
				IServiceProvider divisionSheetProvider = BuildServiceProviderForDivisionSheet(provider, pair.Value, config);
				IDivisionSheetGenerator creator = divisionSheetProvider.GetRequiredService<IDivisionSheetGenerator>();
				PoolPlayInfo pool = await creator.CreateSheet(config);

				// build the provider for the division and add the score entry section to the Shootout sheet
				IServiceProvider shootoutScoreProvider = BuildServiceProviderForShootoutScoreEntry(provider, config, config.DivisionConfigs[pair.Key]);
				IShootoutScoreEntryService scoreEntryService = shootoutScoreProvider.GetRequiredService<IShootoutScoreEntryService>();
				await scoreEntryService.CreateShootoutScoreEntrySection(pool);
			}

			// Hide helper columns on the Shootout sheet after all divisions are complete
			await shootoutSheetService.HideHelperColumns(config);

			logger.LogInformation("All done!");
		}

		private static IServiceProvider BuildSharedServices(AppSettings settings)
		{
			IServiceCollection services = new ServiceCollection();
			services.AddOptions();
			services.AddLogging(builder =>
			{
				builder.ClearProviders();
				builder.AddDebug();
				builder.AddConsole();
			});
			services.AddOAuthChecker(options =>
			{
				options.SecretsPath = settings.ClientSecretsPath;
				options.BrowserArguments = Configuration!.GetSection("OAuth")[nameof(OAuthCheckerOptions.BrowserArguments)];
				options.Scopes.Add(SheetsService.ScopeConstants.DriveFile);
			});
			services.AddSingleton(settings);
			return services.BuildServiceProvider();
		}

		private static IServiceProvider BuildServiceProviderForDocument(AppSettings settings)
		{
			ServiceCollection services = new ServiceCollection();
			services.AddLogging(builder =>
			{
				builder.ClearProviders();
				builder.AddDebug();
				builder.AddConsole();
			});
			services.AddSingleton(settings);
			services.AddSingleton(GoogleCredential!);
			services.AddSingleton<ISheetsClient, SheetsClient>();
			services.AddSingleton<ITeamReaderService, TeamReaderService>();
			services.AddSingleton<IFileReader, FileReader>();
			services.AddSingleton<IShootoutSheetService, ShootoutSheetService>();

			IServiceProvider provider = services.BuildServiceProvider();
			return provider;
		}

		private static IServiceProvider BuildServiceProviderForDivisionSheet(IServiceProvider provider, IEnumerable<Team> divisionTeams, ShootoutSheetConfig config)
		{
			ServiceCollection services = new ServiceCollection();
			services.AddLogging(builder =>
			{
				builder.ClearProviders();
				builder.AddDebug();
				builder.AddConsole();
			});
			services.AddSingleton(provider.GetRequiredService<AppSettings>());
			services.AddSingleton(provider.GetRequiredService<ISheetsClient>());
			services.AddPantherShootoutServices(divisionTeams, config);
			return services.BuildServiceProvider();
		}

		private static IServiceProvider BuildServiceProviderForShootoutScoreEntry(IServiceProvider provider, ShootoutSheetConfig config, DivisionSheetConfig divisionConfig)
		{
			ServiceCollection services = new ServiceCollection();
			services.AddLogging(builder =>
			{
				builder.ClearProviders();
				builder.AddDebug();
				builder.AddConsole();
			});
			services.AddSingleton(provider.GetRequiredService<AppSettings>());
			services.AddSingleton(provider.GetRequiredService<ISheetsClient>());
			services.AddShootoutScoreEntryServices(config, divisionConfig);

			return services.BuildServiceProvider();
		}
	}
}