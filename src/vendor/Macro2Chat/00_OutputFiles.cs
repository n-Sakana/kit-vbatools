using System;
using System.IO;

namespace Macro2Chat
{
    internal static class OutputFiles
    {
        internal static string TemporaryPath(string destination)
        {
            return Path.Combine(Path.GetDirectoryName(destination),
                ".macrostudio-" + Guid.NewGuid().ToString("N") + ".tmp");
        }

        internal static void Publish(string temporaryPath,
            string destination, bool replaceExisting)
        {
            if (!replaceExisting || !File.Exists(destination))
            {
                // The two-argument overload never overwrites a racing file.
                File.Move(temporaryPath, destination);
                return;
            }
            string backup = TemporaryPath(destination);
            try
            {
                // Source, destination and backup are on the same volume.
                // No delete-then-move gap, and no guessed .previous name.
                File.Replace(temporaryPath, destination, backup);
            }
            catch (Exception error)
            {
                if (File.Exists(backup))
                {
                    if (!File.Exists(destination))
                    {
                        try { File.Move(backup, destination); }
                        catch (Exception restoreError)
                        {
                            throw new IOException("The previous output is at " +
                                backup + ". Restore failed: " + restoreError.Message,
                                error);
                        }
                    }
                    else
                    {
                        // Do not delete a destination another application
                        // may have created. Keep the backup for recovery.
                        throw new IOException("The replacement failed. The " +
                            "previous output backup is at " + backup, error);
                    }
                }
                throw;
            }
            try { File.Delete(backup); }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }
}
