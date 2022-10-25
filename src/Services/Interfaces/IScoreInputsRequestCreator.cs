namespace PantherShootoutScoreSheetGenerator.Services
{
	public interface IScoreInputsRequestCreator
	{
		/// <summary>
		/// Creates the <see cref="Request"/> and <see cref="UpdateRequest"/> objects to build the scoring section of the sheet; this is executed once per round of games
		/// </summary>
		/// <param name="info"></param>
		/// <param name="poolTeams"></param>
		/// <param name="roundNum"></param>
		/// <param name="startRowIndex">When passed in, this should be the index of the round label row; when complete, this should be the index for the start of the next round</param>
		/// <returns></returns>
		PoolPlayInfo CreateScoringRequests(PoolPlayInfo info, IEnumerable<Team> poolTeams, int roundNum, ref int startRowIndex);
	}
}