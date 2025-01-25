using System.Reflection;
using Google.Apis.Auth.OAuth2;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using PantherShootoutScoreSheetGenerator.Services;
using GoogleOAuthCliClient;
using GoogleSheetsHelper;
using Google.Apis.Sheets.v4;

namespace PantherShootoutScoreSheetGenerator.ConsoleApp
{
	internal class Program
	{
		private static IConfigurationRoot? Configuration;
		private static GoogleCredential? GoogleCredential;

		static async Task Main(string[] args)
		{
			string? outputPath = Path.GetDirectoryName(new Uri(Assembly.GetExecutingAssembly().Location).LocalPath);
			Configuration = new ConfigurationBuilder()
				.SetBasePath(outputPath)
				.AddJsonFile("appSettings.json")
				.Build();

			await Run();
		}

		static IServiceProvider BuildSharedServices(AppSettings settings)
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

		static IServiceProvider BuildServiceProvider(IServiceProvider provider, IEnumerable<Team> divisionTeams, ShootoutSheetConfig config)
		{
			IServiceCollection services = new ServiceCollection();
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

		static async Task Run()
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
			IServiceCollection services = new ServiceCollection();
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

			// read data and create the team list/shootout sheet
			ITeamReaderService teamReader = provider.GetRequiredService<ITeamReaderService>();
			IDictionary<string, IEnumerable<Team>> teams = teamReader.GetTeams();

			// create the spreadsheet
			ISheetsClient sheetsClient = provider.GetRequiredService<ISheetsClient>();
			await sheetsClient.CreateSpreadsheet($"{DateTime.Today.Year} Panther Shootout scores");
			ILogger<Program> logger = provider.GetRequiredService<ILogger<Program>>();
			logger.LogInformation("Spreadsheet ID {id} created", sheetsClient.SpreadsheetId);

			IShootoutSheetService shootoutSheetService = provider.GetRequiredService<IShootoutSheetService>();
			ShootoutSheetConfig config = await shootoutSheetService.GenerateSheet(teams);

			// loop thru the divisions
			foreach (KeyValuePair<string, IEnumerable<Team>> pair in teams)
			{
				// build the provider for the division and create the sheet
				IServiceProvider divisionSheetProvider = BuildServiceProvider(provider, pair.Value, config);
				IDivisionSheetGenerator creator = divisionSheetProvider.GetRequiredService<IDivisionSheetGenerator>();
				await creator.CreateSheet();
			}

			logger.LogInformation("All done!");
		}
	}
}