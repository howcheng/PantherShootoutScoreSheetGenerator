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
		protected readonly ITiebreakerColumnsRequestCreator _tiebreakersRequestCreator;
		protected readonly ISortedStandingsListRequestCreator _sortedStandingsListRequestCreator;

		public PoolPlayRequestCreator(DivisionSheetConfig config, IScoreSheetHeadersRequestCreator headersCreator, IScoreInputsRequestCreator inputsCreator
			, IStandingsTableRequestCreator tableCreator, ITiebreakerColumnsRequestCreator tiebreakersRequestCreator, ISortedStandingsListRequestCreator sortedStandingsListRequestCreator)
		{
			_config = config;
			_headersRequestCreator = headersCreator;
			_inputsRequestCreator = inputsCreator;
			_standingsTableRequestCreator = tableCreator;
			_tiebreakersRequestCreator = tiebreakersRequestCreator;
			_sortedStandingsListRequestCreator = sortedStandingsListRequestCreator;
		}

		public virtual PoolPlayInfo CreatePoolPlayRequests(PoolPlayInfo info)
		{
			int startRowIndex = 0;
			foreach (IGrouping<string, Team> pool in info.Pools!)
			{
				// pool row and header row (standings)
				info = _headersRequestCreator.CreateHeaderRequests(info, pool.Key, startRowIndex, pool);
				startRowIndex += 1; // startRowIndex is now the row with the round number and standings headers

				// standings table
				info = _standingsTableRequestCreator.CreateStandingsRequests(info, pool, startRowIndex + 1);
				int standingsStartRowNum = startRowIndex + 2;
				int standingsEndRowNum = standingsStartRowNum + _config.TeamsPerPool - 1; // -1 because A3:A6 is 4 rows
				Tuple<int, int> standingsStartAndEnd = new Tuple<int, int>(standingsStartRowNum, standingsEndRowNum);
				info.StandingsStartAndEndRowNums.Add(standingsStartAndEnd);

				// scoring rows
				for (int i = 0; i < _config.NumberOfGameRounds; i++)
				{
					info = _inputsRequestCreator.CreateScoringRequests(info, pool, i + 1, ref startRowIndex);
				}
				info.ScoreEntryStartAndEndRowNums.Add(new Tuple<int, int>(standingsStartRowNum, startRowIndex - 1)); // -1 because of the blank line inserted at the end of each round

				// tiebreaker columns
				info = _tiebreakersRequestCreator.CreateTiebreakerRequests(info, pool, standingsStartRowNum - 1);

				// sorted standings list
				info = _sortedStandingsListRequestCreator.CreateSortedStandingsListRequest(info);
			}

			info.ChampionshipStartRowIndex = startRowIndex;
			return info;
		}
	}
}
