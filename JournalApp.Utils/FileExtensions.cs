using System;
using System.IO;
using System.Reflection;

namespace JournalApp.Utils
{
    public static class FileExtensions
    {
        /// <summary>
        /// Retrieve the size in bytes of the specified file
        /// </summary>
        /// <param name="FilePath"></param>
        /// <returns></returns>
        public static long GetFileSize(this string FilePath)
        {
            if (File.Exists(FilePath))
            {
                return new FileInfo(FilePath).Length;
            }
            return 0;
        }

        /// <summary>
        /// Retrieve directory path of the executing application
        /// </summary>
        /// <returns></returns>
        public static string GetExecutingDirectoryName(this object thing)
        {
            var location = new Uri(Assembly.GetEntryAssembly().GetName().CodeBase);
            return new FileInfo(location.AbsolutePath).Directory.FullName;
        }
    }
}
