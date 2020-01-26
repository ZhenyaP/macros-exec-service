using System.IO;
using System.Threading.Tasks;

namespace MacrosExecService.Helpers
{
    public static class FileHelper
    {
        private const int BufferSize = 4096;

        public static async Task WriteAllBytesAsync(string path, byte[] bytes)
        {

            using (FileStream fileStream = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.Read, BufferSize))
                await fileStream.WriteAsync(bytes, 0, bytes.Length).ConfigureAwait(false);
        }

    }
}
