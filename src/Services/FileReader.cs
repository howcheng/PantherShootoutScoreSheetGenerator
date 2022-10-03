namespace PantherShootoutScoreSheetGenerator.Services
{
	public class FileReader : IFileReader
	{
		public bool DoesFileExist(string path) => File.Exists(path);

		public Stream ReadAsStream(string path)
		{
			FileStream stream = File.OpenRead(path);
			return stream;
		}
	}
}
