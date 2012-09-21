using System;
using System.IO;
using System.Text;

namespace Westwind.Utilities
{
	/// <summary>
	/// wwUtils class which contains a set of common utility classes for 
	/// Formatting strings
	/// Reflection Helpers
	/// Object Serialization
	/// </summary>
	public partial class FileUtils
	{

        /// <summary>
        /// Copies the content of the one stream to another.
        /// Streams must be open and stay open.
        /// </summary>
        public static void CopyStream(Stream source, Stream dest, int bufferSize)
        {
            byte[] buffer = new byte[bufferSize];
            int read;
            while ( (read = source.Read(buffer, 0, buffer.Length)) > 0)
            {
                dest.Write(buffer, 0, read);
            }
        }

        /// <summary>
        /// Detects the byte order mark of a file and returns
        /// an appropriate encoding for the file.
        /// </summary>
        /// <param name="srcFile"></param>
        /// <returns></returns>
        public static Encoding GetFileEncoding(string srcFile)
        {
            // *** Use Default of Encoding.Default (Ansi CodePage)
            Encoding enc = Encoding.Default;

            // *** Detect byte order mark if any - otherwise assume default

            byte[] buffer = new byte[5];
            FileStream file = new FileStream(srcFile, FileMode.Open);
            file.Read(buffer, 0, 5);
            file.Close();

            if (buffer[0] == 0xef && buffer[1] == 0xbb && buffer[2] == 0xbf)
               enc = Encoding.UTF8;
            else if (buffer[0] == 0xfe && buffer[1] == 0xff)
                enc = Encoding.Unicode;
            else if (buffer[0] == 0 && buffer[1] == 0 && buffer[2] == 0xfe && buffer[3] == 0xff)
                enc = Encoding.UTF32;

            else if (buffer[0] == 0x2b && buffer[1] == 0x2f && buffer[2] == 0x76)
                enc = Encoding.UTF7;

            return enc;
        }



        /// <summary>
        /// Opens a stream reader with the appropriate text encoding applied.
        /// </summary>
        /// <param name="srcFile"></param>
        public static StreamReader OpenStreamReaderWithEncoding(string srcFile)
        {
            Encoding enc = GetFileEncoding(srcFile);
            return new StreamReader(srcFile, enc);
        }


		/// <summary>
		/// Returns the full path of a full physical filename
		/// </summary>
		/// <param name="path"></param>
		/// <returns></returns>
		public static string JustPath(string path) 
		{
			FileInfo fi = new FileInfo(path);
			return fi.DirectoryName + "\\";
		}

        /// <summary>
        /// Returns a fully qualified path from a partial or relative
        /// path.
        /// </summary>
        /// <param name="Path"></param>
        public static string GetFullPath(string path)
        {
            if (string.IsNullOrEmpty(path))
                return "";

            return Path.GetFullPath(path);
        }

		/// <summary>
		/// Returns a relative path string from a full path.
		/// </summary>
		/// <param name="FullPath">The path to convert. Can be either a file or a directory</param>
		/// <param name="BasePath">The base path to truncate to and replace</param>
		/// <returns>
		/// Lower case string of the relative path. If path is a directory it's returned without a backslash at the end.
		/// 
		/// Examples of returned values:
		///  .\test.txt, ..\test.txt, ..\..\..\test.txt, ., ..
		/// </returns>
		public static string GetRelativePath(string FullPath, string BasePath ) 
		{
			// *** Start by normalizing paths
			FullPath = FullPath.ToLower();
			BasePath = BasePath.ToLower();

			if ( BasePath.EndsWith("\\") ) 
				BasePath = BasePath.Substring(0,BasePath.Length-1);
			if ( FullPath.EndsWith("\\") ) 
				FullPath = FullPath.Substring(0,FullPath.Length-1);

			// *** First check for full path
			if ( (FullPath+"\\").IndexOf(BasePath + "\\") > -1) 
				return  FullPath.Replace(BasePath,".");

			// *** Now parse backwards
			string BackDirs = "";
			string PartialPath = BasePath;
			int Index = PartialPath.LastIndexOf("\\");
			while (Index > 0) 
			{
				// *** Strip path step string to last backslash
				PartialPath = PartialPath.Substring(0,Index );
			
				// *** Add another step backwards to our pass replacement
				BackDirs = BackDirs + "..\\" ;

				// *** Check for a matching path
				if ( FullPath.IndexOf(PartialPath) > -1 ) 
				{
					if ( FullPath == PartialPath )
						// *** We're dealing with a full Directory match and need to replace it all
						return FullPath.Replace(PartialPath,BackDirs.Substring(0,BackDirs.Length-1) );
					else
						// *** We're dealing with a file or a start path
						return FullPath.Replace(PartialPath+ (FullPath == PartialPath ?  "" : "\\"),BackDirs);
				}
				Index = PartialPath.LastIndexOf("\\",PartialPath.Length-1);
			}

			return FullPath;
		}

        /// <summary>
        /// Deletes files based on a file spec and a given timeout.
        /// This routine is useful for cleaning up temp files in 
        /// Web applications.
        /// </summary>
        /// <param name="filespec">A DOS filespec that includes path and/or wildcards to select files</param>
        /// <param name="seconds">The timeout - if files are older than this timeout they are deleted</param>
        public static void DeleteFiles(string filespec,int seconds)
        {
            string path = Path.GetDirectoryName(filespec);
            string spec = Path.GetFileName(filespec);
            string[] files = Directory.GetFiles(path,spec);
            foreach(string file in files)
            {
                try
                {
                    if (File.GetLastWriteTimeUtc(file) < DateTime.UtcNow.AddSeconds(seconds * -1))
                        File.Delete(file);
                }
                catch {}  // ignore locked files
            }
        }
            
		
    }

}