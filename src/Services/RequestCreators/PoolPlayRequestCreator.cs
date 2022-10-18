using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	/// <summary>
	/// 
	/// </summary>
	public class PoolPlayRequestCreator : IPoolPlayRequestCreator
	{
		protected readonly DivisionSheetConfig _config;
		protected readonly IScoreSheetHeadersRequestCreator _headersRequestCreator;
		protected readonly IScoreInputsRequestCreator _inputsRequestCreator;
		protected readonly IStandingsTableRequestCreator _standingsTableRequestCreator;

		public PoolPlayRequestCreator(DivisionSheetConfig config, IScoreSheetHeadersRequestCreator headersCreator, IScoreInputsRequestCreator inputsCreator, IStandingsTableRequestCreator tableCreator)
		{
			_config = config;
			_headersRequestCreator = headersCreator;
			_inputsRequestCreator = inputsCreator;
			_standingsTableRequestCreator = tableCreator;
		}

		public virtual Task<PoolPlayInfo> CreatePoolPlayRequests(PoolPlayInfo info)
		{
			int startRowIndex = 0;
			foreach (IGrouping<string, Team> pool in info.Pools!)
			{
				// pool row and header row (standings)
				info = _headersRequestCreator.CreateHeaderRequests(info, pool.Key, startRowIndex, pool);
				startRowIndex += 1; // startRowIndex is now the row with the round number and standings headers

				// standings table
				info = _standingsTableRequestCreator.CreateStandingsRequests(_config, info, pool, startRowIndex + 1);
				int standingsStartRowNum = startRowIndex + 2;
				info.StandingsStartAndEndRowNums.Add(new Tuple<int, int>(standingsStartRowNum, standingsStartRowNum + _config.TeamsPerPool - 1)); // -1 because A3:A6 is 4 rows

				// scoring rows
				for (int i = 0; i < _config.NumberOfRounds; i++)
				{
					info = _inputsRequestCreator.CreateScoringRequests(_config, info, pool, i + 1, ref startRowIndex);
				}
			}

			info.ChampionshipStartRowIndex = startRowIndex;
			return Task.FromResult(info);
		}
	}
}
