using GoogleSheetsHelper;
using StandingsGoogleSheetsHelper;

namespace PantherShootoutScoreSheetGenerator.Services
{
	public class ChampionshipRequestCreator6Teams : StandardChampionshipRequestCreator
	{
		private readonly PsoFormulaGenerator _formGen;

		public ChampionshipRequestCreator6Teams(DivisionSheetConfig config, FormulaGenerator formGen)
			: base(config, formGen)
		{
			_formGen = (PsoFormulaGenerator)formGen;
		}

		/// <summary>
		/// In a 6-team division, the two 3rd-place teams play each other for the consolation, then top 2 teams from each pool do semifinal and final rounds
		/// </summary>
		/// <param name="info"></param>
		/// <returns></returns>
		public override ChampionshipInfo CreateChampionshipRequests(PoolPlayInfo info)
		{
			ChampionshipInfo ret = new ChampionshipInfo(info);

			int startRowIndex = info.ChampionshipStartRowIndex;
			List<UpdateRequest> updateRequests = new List<UpdateRequest>();

			string pool1 = ret.Pools!.First().First().PoolName;
			string pool2 = ret.Pools!.Last().First().PoolName;

			// finals label row
			UpdateRequest labelRequest = Utilities.CreateHeaderLabelRowRequest(_divisionName, startRowIndex, _helper.StandingsTableColumns.Count - 1, "FINALS", 4, cell => cell.SetHeaderCellFormatting());
			updateRequests.Add(labelRequest);
			startRowIndex += 1; // this is now the subheader label for the consolation game

			// consolation: 3rd place teams from each pool
			UpdateRequest subheaderRequest = Utilities.CreateHeaderLabelRowRequest(_divisionName, startRowIndex, _helper.HeaderRowColumns.Count - 1, $"CONSOLATION: 3rd from Pool {pool1} vs 3rd from Pool {pool2}", 0, cell => cell.SetSubheaderCellFormatting());
			updateRequests.Add(subheaderRequest);
			startRowIndex += 1; // this is now the scores row for the consolation game
			string homeFormula, awayFormula;
			homeFormula = CreateSemifinalFormula(3, info.StandingsStartAndEndRowNums.First().Item1, info.StandingsStartAndEndRowNums.First().Item2);
			awayFormula = CreateSemifinalFormula(3, info.StandingsStartAndEndRowNums.Last().Item1, info.StandingsStartAndEndRowNums.Last().Item2);
			updateRequests.Add(CreateChampionshipGameRequests(startRowIndex, homeFormula, awayFormula));
			ret.ConsolationGameRowNum = startRowIndex + 1; // Save the consolation game row number
			startRowIndex += 1; // this is now the subheader label for the semifinals

			// semifinals: 1st place from pool 1 vs 2nd place from pool 2 (and vice versa)
			updateRequests.Add(CreateChampionshipHeaderRow(startRowIndex, "SEMIFINALS: 1st place in each pool vs 2nd place from the other pool"));
			startRowIndex += 1; // this is now the 1st scores row for the semifinals
			int startRowNum = startRowIndex + 1;
			int semi1RowNum = startRowNum;

			CreateSemifinalFormulas(info, true, out homeFormula, out awayFormula);
			updateRequests.Add(CreateChampionshipGameRequests(startRowIndex, homeFormula, awayFormula));
			ret.Semifinal1RowNum = semi1RowNum;
			startRowIndex += 1; // this is now the 2nd scores row for the semifinals
			int semi2RowNum = startRowNum += 1;

			CreateSemifinalFormulas(info, false, out homeFormula, out awayFormula);
			updateRequests.Add(CreateChampionshipGameRequests(startRowIndex, homeFormula, awayFormula));
			ret.Semifinal2RowNum = semi2RowNum;
			startRowIndex += 1; // this is now the subheader label for the 3rd-place game
			startRowNum += 1;

			// 3rd-place game
			updateRequests.Add(CreateChampionshipHeaderRow(startRowIndex, "3RD-PLACE GAME"));
			startRowIndex += 1; // this is now the scores row for the 3rd-place game
			startRowNum += 1;
			CreateFinalFormulas(semi1RowNum, semi2RowNum, false, out homeFormula, out awayFormula);
			updateRequests.Add(CreateChampionshipGameRequests(startRowIndex, homeFormula, awayFormula));
			ret.ThirdPlaceGameRowNum = startRowNum; // This is the actual 3rd place game (no shootout)
			startRowIndex += 1; // this is now the subheader label for the championship
			startRowNum += 1;

			// final
			updateRequests.Add(CreateChampionshipHeaderRow(startRowIndex, "FINAL"));
			startRowIndex += 1; // this is now the scores row for the championship
			startRowNum += 1;
			CreateFinalFormulas(semi1RowNum, semi2RowNum, true, out homeFormula, out awayFormula);
			updateRequests.Add(CreateChampionshipGameRequests(startRowIndex, homeFormula, awayFormula));
			ret.ChampionshipGameRowNum = startRowNum;

			ret.UpdateValuesRequests = updateRequests;
			return ret;
		}

		private void CreateSemifinalFormulas(PoolPlayInfo info, bool firstSemi, out string homeFormula, out string awayFormula)
		{
			// for first semifinal, it's 1st from pool (H) vs 2nd from pool 2 (A), then 2nd from pool 1 (H) vs 1st from pool 2 (A)
			int pool1StartRow = info.StandingsStartAndEndRowNums.First().Item1;
			int pool1EndRow = info.StandingsStartAndEndRowNums.First().Item2;
			int pool2StartRow = info.StandingsStartAndEndRowNums.Last().Item1;
			int pool2EndRow = info.StandingsStartAndEndRowNums.Last().Item2;
			homeFormula = CreateSemifinalFormula(firstSemi ? 1 : 2, pool1StartRow, pool1EndRow);
			awayFormula = CreateSemifinalFormula(firstSemi ? 2 : 1, pool2StartRow, pool2EndRow);
		}

		private string CreateSemifinalFormula(int rank, int standingsStartRowNum, int standingsEndRowNum) 
			=> _formGen.GetTeamNameFromStandingsTableByRankFormula(rank, _config.TotalPoolPlayGames, standingsStartRowNum, standingsEndRowNum);

		private void CreateFinalFormulas(int semi1RowNum, int semi2RowNum, bool isChampionship, out string homeFormula, out string awayFormula)
		{
			// for third-place game, loser of semifinal 1 (H) vs loser of semifinal 2 (A)
			// for championship, winner of semifinal 2 (H) vs winner of semifinal 1 (A)
			homeFormula = CreateFinalFormula(isChampionship ? semi1RowNum : semi2RowNum, isChampionship);
			awayFormula = CreateFinalFormula(isChampionship ? semi2RowNum : semi1RowNum, isChampionship);
		}

		private string CreateFinalFormula(int rowNum, bool isChampionship)
		{
			// =IFNA(IFS(B25>C25,A25,B25<C25,D25), ""): home team for final
			// for away team, use the other row number
			// for 3rd-place game, switch the < and > operators
			char comparison1 = isChampionship ? '>' : '<';
			char comparison2 = isChampionship ? '<' : '>';
			return string.Format("=IFNA(IFS({4}{0}{1}{5}{0},{3}{0},{4}{0}{2}{5}{0},{6}{0}), \"\")",
				rowNum,
				comparison1,
				comparison2,
				_helper.HomeTeamColumnName,
				_helper.HomeGoalsColumnName,
				_helper.AwayGoalsColumnName,
				_helper.AwayTeamColumnName);
		}
	}
}
