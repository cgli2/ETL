using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using SharpCompress.Reader;
using SharpCompress.Archive;
using SharpCompress.Common;

namespace ExcelTool
{
   public class ZipHelper
    {
        public static void UnRar(String file,String saveFolder)
        {
            using (Stream stream = File.OpenRead(file))
            {
                using (var reader = ReaderFactory.Open(stream))
                {
                    while (reader.MoveToNextEntry())
                    {
                        if (!reader.Entry.IsDirectory)
                        {
                            Console.WriteLine(reader.Entry.FilePath);
                            reader.WriteEntryToDirectory(saveFolder, ExtractOptions.ExtractFullPath | ExtractOptions.Overwrite);
                        }
                    }
                }
            }
        }

        public static void UnZip(String file, String saveFolder)
        {
            var archive = ArchiveFactory.Open(file);
           
            foreach (var entry in archive.Entries)
            {
                if (!entry.IsDirectory)
                {
                    Console.WriteLine(entry.FilePath);
                    entry.WriteToDirectory(saveFolder, ExtractOptions.ExtractFullPath | ExtractOptions.Overwrite);
                }
            }

        }
    }
}
