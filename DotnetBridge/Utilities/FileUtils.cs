using System;
using System.IO;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace Westwind.WebConnection
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
        /// path. Also fixes up case of the entire path **if the file exists**
        /// </summary>
        /// <param name="Path"></param>
        public static string GetFullPath(string path)
        {
            if (string.IsNullOrEmpty(path))
                return "";

            path = Path.GetFullPath(path);

            // Ensure we fix case
            if (File.Exists(path))
            {
                StringBuilder builder = new StringBuilder(255);

                // names with long extension can cause the short name to be actually larger than
                // the long name.
                GetShortPathName(path, builder, builder.Capacity);
                string shortPath = builder.ToString();

                GetLongPathName(shortPath, builder, builder.Capacity);
                path = builder.ToString();
            }

            return path;
        }
        

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        static extern uint GetLongPathName(string ShortPath, StringBuilder sb, int buffer);

        [DllImport("kernel32.dll")]
        static extern uint GetShortPathName(string longpath, StringBuilder sb, int buffer);
        
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
        /// <summary>
        /// Returns a relative path string from a full path based on a base path
        /// provided.
        /// </summary>
        /// <param name="fullPath">The path to convert. Can be either a file or a directory</param>
        /// <param name="basePath">The base path on which relative processing is based. Should be a directory.</param>
        /// <returns>
        /// String of the relative path.
        /// 
        /// Examples of returned values:
        ///  test.txt, ..\test.txt, ..\..\..\test.txt, ., .., subdir\test.txt
        /// </returns>
        public static string GetRelativePath(string fullPath, string basePath)
        {
            if (string.IsNullOrEmpty(fullPath))
                return fullPath;

            // get the full path from any partial paths
            fullPath = Path.GetFullPath(fullPath);

            // Ensure we fix case
            if (File.Exists(fullPath))
            {
                StringBuilder builder = new StringBuilder(255);

                // names with long extension can cause the short name to be actually larger than
                // the long name.
                GetShortPathName(fullPath, builder, builder.Capacity);
                string path = builder.ToString();

                uint result = GetLongPathName(path, builder, builder.Capacity);
                fullPath = builder.ToString();
            }

            // ForceBasePath to a path
            if (!basePath.EndsWith(value: "\\"))
                basePath += "\\";

#pragma warning disable CS0618
            Uri baseUri = new Uri(uriString: basePath, dontEscape: true);
            Uri fullUri = new Uri(uriString: fullPath, dontEscape: true);
#pragma warning restore CS0618

            Uri relativeUri = baseUri.MakeRelativeUri(uri: fullUri);

            // Uri's use forward slashes so convert back to backward slahes
            return relativeUri.ToString().Replace(oldValue: "/", newValue: "\\");
        }


        /// <summary>
        /// Returns a filename for a file to save
        /// </summary>
        /// <param name="folder"></param>
        /// <param name="caption"></param>
        /// <param name="title"></param>
        /// <param name="defaultExtension">a default extension like png, txt, md etc.</param>
        /// <param name="extensionFilter">"|Png Image (.png)|*.png|JPEG Image (.jpeg)|*.jpeg|Gif Image (.gif)|*.gif|Tiff Image (.tiff)|*.tiff|All Files (*.*)|*.*";</param>
        /// <param name="promptOverwrite">If true prompts if file exists</param>
        /// <returns>Filename or null</returns>
        public static string SaveFileDialog(string folder, string title, 
            string defaultExtension, string extensionFilter, bool promptOverwrite)
        {
            string file = null;
            if (Path.HasExtension(folder))
            {
                file = folder;
                folder = Path.GetDirectoryName(folder);                
            }

            var dialog = new SaveFileDialog();
            dialog.RestoreDirectory = true;
            dialog.InitialDirectory = folder;
            if (!string.IsNullOrEmpty(file))
                dialog.FileName = file;            
            dialog.Title = title;
            dialog.AddExtension = true;
            dialog.DefaultExt = defaultExtension;
            dialog.Filter = extensionFilter;
            dialog.OverwritePrompt = promptOverwrite;

            var result = dialog.ShowDialog();
            if (result != DialogResult.OK)
                return null;

            return dialog.FileName;
        }


        /// <summary>
        /// Returns a filename for a file to open
        /// </summary>
        /// <param name="folder"></param>
        /// <param name="caption"></param>
        /// <param name="title"></param>        
        /// <param name="extensionFilter">"|Png Image (.png)|*.png|JPEG Image (.jpeg)|*.jpeg|Gif Image (.gif)|*.gif|Tiff Image (.tiff)|*.tiff|All Files (*.*)|*.*";</param>
        /// <param name="checkIfFileExists">If true requires that the selected file exists. Otherwise you can type a filename.</param>
        /// <returns>Filename or null</returns>
        public static string OpenFileDialog(string folder, string title,
            string extensionFilter, bool checkIfFileExists)
        {
            string file = null;
            if (Path.HasExtension(folder))
            {
                file = folder;
                folder = Path.GetDirectoryName(folder);
            }

            var dialog = new OpenFileDialog();
            dialog.RestoreDirectory = true;
            dialog.InitialDirectory = folder;
            if (!string.IsNullOrEmpty(file))
                dialog.FileName = file;
            dialog.Title = title;
            dialog.AddExtension = true;            
            dialog.Filter = extensionFilter;
            dialog.CheckFileExists = checkIfFileExists;

            var result = dialog.ShowDialog();
            if (result != DialogResult.OK)
                return null;

            return dialog.FileName;
        }

        /// <summary>
        /// Displays a folder browser dialog box
        /// </summary>
        /// <param name="startFolder"></param>
        /// <param name="description"></param>
        /// <returns></returns>
        public static string OpenFolderDialog(string startFolder, string description)
        {
            if (description == string.Empty)
                description = null;
            if (string.IsNullOrEmpty(startFolder))
                startFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            var dialog = new FolderBrowserDialog()
            {
                SelectedPath = startFolder,
                Description = description,
                ShowNewFolderButton = true
            };
            var dialogResult = dialog.ShowDialog();

            if (dialogResult != DialogResult.OK)
                return null;

            return dialog.SelectedPath;
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

        /// <summary>
        /// Zips a folder 
        /// </summary>
        /// <param name="outputZipFile"></param>
        /// <param name="folder"></param>
        /// <returns></returns>
        public static bool ZipFolder(string outputZipFile, string folder, bool fast)
        {
            try
            {
                if (File.Exists(outputZipFile))
                    File.Delete(outputZipFile);

                ZipFile.CreateFromDirectory(folder, 
                        outputZipFile, 
                        fast ? CompressionLevel.Fastest :CompressionLevel.Optimal, 
                        false);
            }
            catch
            {
                return false;
            }

            return true;
        }

        /// <summary>
        /// Unzips a zip file to a destination folder.
        /// </summary>
        /// <param name="zipFile"></param>
        /// <param name="folder"></param>
        /// <returns></returns>
        public static bool UnzipFolder(string zipFile, string folder)
        {
            try
            {
                ZipFile.ExtractToDirectory(zipFile, folder);
            }
            catch
            {
                return false;
            }

            return true;
        }

    }

}