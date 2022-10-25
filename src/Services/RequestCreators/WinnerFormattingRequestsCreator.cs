using Google.Apis.Sheets.v4.Data;
using GoogleSheetsHelper;
using Humanizer;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public class WinnerFormattingRequestsCreator : IWinnerFormattingRequestsCreator
	{
		protected readonly DivisionSheetConfig _config;
		protected readonly PsoDivisionSheetHelper _helper;

		public WinnerFormattingRequestsCreator(DivisionSheetConfig config, PsoDivisionSheetHelper helper)
		{
			_config = config;
			_helper = helper;
		}

		/// <summary>
		/// Creates the <see cref="Request"/> and <see cref="UpdateRequest"/> objects to display the division winners
		/// </summary>
		/// <param name="info"></param>
		/// <returns>The requests to send to Google Sheets</returns>
		/// <remarks>
		/// Even though <see cref="ChampionshipInfo"/> inherits from <see cref="SheetRequests"/> it will already have its own requests as a result of <see cref="CreateChampionshipRequests(PoolPlayInfo)"/>
		/// so that's why we are returning a new object
		/// </remarks>
		public SheetRequests CreateWinnerFormattingRequests(ChampionshipInfo info)
		{
			SheetRequests ret = new SheetRequests();
			List<Request> requests = new List<Request>();

			for (int i = 0; i < 4; i++)
			{
				Request formatRequest = CreateWinnerFormattingRequest(i + 1, info);
				requests.Add(formatRequest);
			}

			// legend -- we put this after the first standings table
			int legendStartRowIdx = info.StandingsStartAndEndRowNums.First().Item2 + 1;
			UpdateRequest legendRequest = new UpdateRequest(_config.DivisionName)
			{
				RowStart = legendStartRowIdx,
				ColumnStart = _helper.GetColumnIndexByHeader(Constants.HDR_TEAM_NAME),
			};

			for (int i = 1; i <= 4; i++)
			{
				GoogleSheetRow legendRow = new GoogleSheetRow();
				legendRow.Add(new GoogleSheetCell(i.Ordinalize()));
				legendRow.Add(new GoogleSheetCell(string.Empty).SetBackgroundColor(Colors.GetColorForRank(i)));
				legendRequest.Rows.Add(legendRow);
			}

			// right-align those cells
			Request alignRequest = new Request
			{
				RepeatCell = new RepeatCellRequest
				{
					Range = new GridRange
					{
						SheetId = _config.SheetId,
						StartRowIndex = legendStartRowIdx,
						StartColumnIndex = _helper.GetColumnIndexByHeader(Constants.HDR_TEAM_NAME),
						EndRowIndex = legendStartRowIdx + 4,
						EndColumnIndex = _helper.GetColumnIndexByHeader(Constants.HDR_TEAM_NAME) + 1,
					},
					Cell = new CellData
					{
						UserEnteredFormat = new CellFormat
						{
							HorizontalAlignment = "RIGHT",
						},
					},
					Fields = "userEnteredFormat(horizontalAlignment)",
				}
			};
			requests.Add(alignRequest);

			ret.UpdateValuesRequests.Add(legendRequest);
			ret.UpdateSheetRequests = requests;
			return ret;
		}

		/// <summary>
		/// Creates the <see cref="Request"/> object to do conditional formatting in the standings table to highlight the winners (1st to 4th place)
		/// </summary>
		/// <param name="rank">This determines what the background color will be</param>
		/// <param name="championshipInfo"></param>
		/// <returns></returns>
		protected virtual Request CreateWinnerFormattingRequest(int rank, ChampionshipInfo championshipInfo)
		{
			// =OR(AND($A$30=$E3, $B$30>$C$30), AND($D$30=$E3, $B$30<$C$30)) -- game winner
			// for loser, switch the positions of the < and >

			bool isChampionship = rank <= 2;
			int gameRowNum = isChampionship ? championshipInfo.ChampionshipGameRowNum : championshipInfo.ThirdPlaceGameRowNum;
			int standingsStartRowNum = championshipInfo.FirstStandingsRowNum;
			char comparison1 = (rank % 2 == 1) ? '>' : '<';
			char comparison2 = (rank % 2 == 1) ? '<' : '>';
			string formula = string.Format("=OR(AND(${4}${0}=${8}{1}, ${5}${0}{2}${6}${0}), AND(${7}${0}=${8}{1}, ${5}${0}{3}${6}${0}))",
				gameRowNum,
				standingsStartRowNum,
				comparison1,
				comparison2,
				_helper.HomeTeamColumnName,
				_helper.HomeGoalsColumnName,
				_helper.AwayGoalsColumnName,
				_helper.AwayTeamColumnName,
				_helper.TeamNameColumnName
			);

			return CreateWinnerConditionalFormattingRequest(rank, formula, championshipInfo.StandingsStartAndEndIndices);
		}

		protected Request CreateWinnerConditionalFormattingRequest(int rank, string formula, IEnumerable<Tuple<int, int>> startAndEndRowIndices)
		{
			Color color = Colors.GetColorForRank(rank);
			Request request = new Request
			{
				AddConditionalFormatRule = new AddConditionalFormatRuleRequest
				{
					Rule = new ConditionalFormatRule
					{
						Ranges = startAndEndRowIndices.Select(pair => new GridRange
						{
							SheetId = _config.SheetId,
							StartRowIndex = pair.Item1,
							StartColumnIndex = _helper.GetColumnIndexByHeader(Constants.HDR_TEAM_NAME),
							EndRowIndex = pair.Item2 + 1,
							EndColumnIndex = _helper.GetColumnIndexByHeader(Constants.HDR_CALC_RANK),
						}).ToList(),
						BooleanRule = new BooleanRule
						{
							Condition = new BooleanCondition
							{
								Type = "CUSTOM_FORMULA",
								Values = new List<ConditionValue>
								{
									new ConditionValue
									{
										UserEnteredValue = formula
									},
								},
							},
							Format = new CellFormat
							{
								BackgroundColor = color,
								TextFormat = new TextFormat
								{
									Bold = true,
								}
							},
						},
					}
				},
			};
			return request;
		}
	}
}
