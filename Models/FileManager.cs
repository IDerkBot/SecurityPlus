using System;
using System.IO;
using Newtonsoft.Json;

namespace SecurityPlus.Models
{
	internal class FileManager
	{
		public static string GetPathSettings()
		{
			var path = $@"{GetFolderAppData()}settings.json";
			if (File.Exists(path)) return path;
			File.Create(path).Dispose();
			var settings = new Settings();
			File.WriteAllText(path, JsonConvert.SerializeObject(settings));
			return path;
		}
		public static Settings GetSettings() { return JsonConvert.DeserializeObject<Settings>(File.ReadAllText(GetPathSettings())); }
		public static void SetSettings(Settings settings) { File.WriteAllText(GetPathSettings(), JsonConvert.SerializeObject(settings)); }
		public static string GetFolderAppData()
		{
			var appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
			var path = $@"{appDataPath}\SecurityTest\";
			if (!Directory.Exists(path))
				Directory.CreateDirectory($@"{appDataPath}\SecurityTest\");
			return path;
		}
  }
}
