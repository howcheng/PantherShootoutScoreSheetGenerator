using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	/// <summary>
	/// 
	/// </summary>
	public class PoolPlayRequestCreator : IPoolPlayRequestCreator
	{
		protected readonly DivisionSheetConfig _config;
		protected readonly string _divisionName;
		protected readonly PsoDivisionSheetHelper _helper;
		protected readonly IStandingsRequestCreatorFactory _requestCreatorFactory;

		public PoolPlayRequestCreator(DivisionSheetConfig config, string divisionName, PsoDivisionSheetHelper helper, IStandingsRequestCreatorFactory factory)
		{
			_config = config;
			_divisionName = divisionName;
			_helper = helper;
			_requestCreatorFactory = factory;
		}

		public virtual Task<PoolPlayInfo> CreatePoolPlayRequests(PoolPlayInfo info)
		{
			int startRowIndex = 0;
			ScoreSheetHeadersRequestCreator headersCreator = new ScoreSheetHeadersRequestCreator(_divisionName, _helper);
			ScoreInputsRequestCreator inputsCreator = new ScoreInputsRequestCreator(_divisionName, _helper, _requestCreatorFactory);
			foreach (IGrouping<string, Team> pool in info.Pools!)
			{
				// pool row and header row (standings)
				info = headersCreator.CreateHeaderRequests(info, pool.Key, startRowIndex, pool); // after this, startRowIndex is the row with the round number and standings headers
				startRowIndex += 1;

				// standings table
				StandingsTableRequestCreator standingsRequestCreator = new StandingsTableRequestCreator(_helper, _requestCreatorFactory);
				info = standingsRequestCreator.CreateStandingsRequests(_config, info, pool, startRowIndex + 1);
				int standingsStartRowNum = startRowIndex + 2;
				info.StandingsStartAndEndRowNums.Add(new Tuple<int, int>(standingsStartRowNum, standingsStartRowNum + _config.TeamsPerPool - 1)); // -1 because A3:A7 is 4 rows

				// scoring rows
				for (int i = 0; i < _config.NumberOfRounds; i++)
				{
					info = inputsCreator.CreateScoringRequests(_config, info, pool, i + 1, ref startRowIndex);
				}
			}

			info.ChampionshipStartRowIndex = startRowIndex;
			return Task.FromResult(info);
		}
	}
}
