namespace PantherShootoutScoreSheetGenerator.Services
{
	public interface IScoreInputsRequestCreator
	{
		PoolPlayInfo CreateScoringRequests(DivisionSheetConfig config, PoolPlayInfo info, IEnumerable<Team> poolTeams, int roundNum, ref int startRowIndex);
	}
}