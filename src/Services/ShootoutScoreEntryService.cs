using GoogleSheetsHelper;
using Microsoft.Extensions.Logging;

namespace PantherShootoutScoreSheetGenerator.Services
{
	/// <summary>
	/// Creates the score entry sections, tiebreaker columns, and sorted standings lists on the Shootout sheet.
	/// </summary>
	public sealed class ShootoutScoreEntryService : IShootoutScoreEntryService
	{
		private readonly ISheetsClient _sheetsClient;
		private readonly ILogger<ShootoutScoreEntryService> _logger;
		private readonly IShootoutScoreEntryRequestsCreator _requestsCreator;
		private readonly ShootoutSheetConfig _shootoutSheetConfig;
		private readonly DivisionSheetConfig _divisionSheetConfig;

		public ShootoutScoreEntryService(ISheetsClient sheetsClient, ILogger<ShootoutScoreEntryService> logger, IShootoutScoreEntryRequestsCreator requestsCreator
			, ShootoutSheetConfig shootoutSheetConfig, DivisionSheetConfig divisionSheetConfig)
		{
			_sheetsClient = sheetsClient;
			_logger = logger;
			_requestsCreator = requestsCreator;
			_shootoutSheetConfig = shootoutSheetConfig;
			_divisionSheetConfig = divisionSheetConfig;
		}

		public async Task CreateShootoutScoreEntrySection(PoolPlayInfo info)
		{
			_logger.LogInformation("Creating shootout score entry section for {divisionName}", _divisionSheetConfig.DivisionName);

			int startRowIdx = _shootoutSheetConfig.ShootoutStartAndEndRows[_divisionSheetConfig.DivisionName].Item1 - 2; // -1 to convert to idx, -1 more for header row
			SheetRequests requests = _requestsCreator.CreateScoreEntryRequests(info, startRowIdx);

			await _sheetsClient.Update(requests.UpdateValuesRequests);
			await _sheetsClient.ExecuteRequests(requests.UpdateSheetRequests);
		}
	}
}
