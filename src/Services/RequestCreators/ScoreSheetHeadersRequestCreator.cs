﻿using Google.Apis.Sheets.v4.Data;
using GoogleSheetsHelper;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public class ScoreSheetHeadersRequestCreator
	{
		private readonly string _divisionName;
		private readonly PsoDivisionSheetHelper _helper;

		public ScoreSheetHeadersRequestCreator(string divisionName, PsoDivisionSheetHelper helper)
		{
			_divisionName = divisionName;
			_helper = helper;
		}

		public virtual PoolPlayInfo CreateHeaderRequests(PoolPlayInfo info, string poolName, int startRowIndex, IEnumerable<Team> poolTeams)
		{
			IList<UpdateRequest> updateRequests = info.UpdateValuesRequests;

			// pool name row
			int poolEndColIndex = _helper.GetColumnIndexByHeader(_helper.StandingsTableColumns.Last());
			UpdateRequest poolNameRequest = new UpdateRequest(_divisionName)
			{
				RowStart = startRowIndex,
			};
			GoogleSheetRow poolNameRow = Utilities.CreateHeaderLabelRow(poolEndColIndex, $"POOL {poolName}", 4, cell => cell.SetHeaderCellFormatting());
			poolNameRequest.Rows.Add(poolNameRow);
			updateRequests.Add(poolNameRequest);

			// standings headers
			int headerRowIndex = startRowIndex + 1;
			UpdateRequest standingsHeaderRequest = new UpdateRequest(_divisionName)
			{
				RowStart = headerRowIndex,
				ColumnStart = _helper.GetColumnIndexByHeader(Constants.HDR_TEAM_NAME),
			};
			standingsHeaderRequest.Rows.Add(_helper.CreateHeaderRow(_helper.StandingsTableColumns, cell => cell.SetBoldText()));
			updateRequests.Add(standingsHeaderRequest);

			// tiebreaker headers
			UpdateRequest tiebreakerHeaderRequest = new UpdateRequest(_divisionName)
			{
				RowStart = headerRowIndex,
				ColumnStart = poolEndColIndex + 1,
			};
			GoogleSheetRow tiebreakerHeaderRow = new GoogleSheetRow();
			tiebreakerHeaderRow.AddRange(poolTeams.Select(x => new GoogleSheetCell() { FormulaValue = $"=\"vs \" & LEFT({ShootoutConstants.SHOOTOUT_SHEET_NAME}!{x.TeamSheetCell}, 2)" }));
			tiebreakerHeaderRequest.Rows.Add(tiebreakerHeaderRow);
			updateRequests.Add(tiebreakerHeaderRequest);

			// note for the manual tiebreaker column
			int tiebreakerColIdx = _helper.GetColumnIndexByHeader(Constants.HDR_TIEBREAKER);
			Request noteRequest = new Request
			{
				UpdateCells = new UpdateCellsRequest
				{
					Range = new GridRange
					{
						StartRowIndex = headerRowIndex,
						EndRowIndex = headerRowIndex + 1,
						StartColumnIndex = tiebreakerColIdx,
						EndColumnIndex = tiebreakerColIdx + 1,
					},
					Rows = new List<RowData>
					{
						new RowData
						{
							Values = new List<CellData>
							{
								new CellData { Note = "This column is in case all of the tiebreakers are equal and it comes down to a coin toss. Enter 1 for the winner." }
							}
						}
					}
				}
			};
			info.UpdateSheetRequests.Add(noteRequest);

			return info;
		}
	}
}
