
namespace TableParser.Output
{
	public abstract class BaseOutputWriter
	{
		protected readonly string OutputPath;
		protected readonly string FileName;

		public BaseOutputWriter(string outputPath, string fileName)
		{
			OutputPath = outputPath;
			FileName = fileName;
		}

		protected void CreateDirectoryIfNeeded()
		{
			if (!Directory.Exists(OutputPath))
			{
				Directory.CreateDirectory(OutputPath);
			}
		}

		protected virtual string GetFileName()
		{
			var now = DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss");
			return FileName + now + ".csv";
		}

		protected virtual string GetFilePath()
		{
			return OutputPath + Path.DirectorySeparatorChar + GetFileName();
		}
	}
}


