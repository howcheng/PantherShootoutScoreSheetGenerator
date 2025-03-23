namespace PantherShootoutScoreSheetGenerator.Services
{
	/// <summary>
	/// Creates the requests to build the score entry fields, the standings table, and the championship enty fields on the division sheet
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
				int standingsStartRowNum = info.StandingsStartAndEndRowNums.Last().Item1; // this was set in the standings table request creator

				List<Tuple<int, int>> scoringStartAndEndRows = new(_config.NumberOfGameRounds);
				info.ScoreEntryStartAndEndRowNums.Add(pool.Key, scoringStartAndEndRows);
				// scoring rows
				for (int i = 0; i < _config.NumberOfGameRounds; i++)
				{
					int scoringStartRow = startRowIndex + 2; // +1 for label row, +1 to convert from index
					int scoringEndRow = scoringStartRow + _config.GamesPerRound - 1;
					info = _inputsRequestCreator.CreateScoringRequests(info, pool, i + 1, ref startRowIndex);
					Tuple<int, int> scoringStartAndEnd = new(scoringStartRow, scoringEndRow);
					scoringStartAndEndRows.Add(scoringStartAndEnd);
				}
				
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
