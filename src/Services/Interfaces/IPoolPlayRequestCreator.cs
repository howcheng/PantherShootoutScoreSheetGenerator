namespace PantherShootoutScoreSheetGenerator.Services
{
	public interface IPoolPlayRequestCreator
	{
		Task<PoolPlayInfo> CreatePoolPlayRequests(PoolPlayInfo info);
	}
}