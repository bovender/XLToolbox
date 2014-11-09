using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Security.Cryptography;

namespace Bovender
{
    public static class FileHelpers
    {
        /// <summary>
        /// Computes the Sha1 hash of a given file.
        /// </summary>
        /// <param name="file">File to compute the Sha1 for.</param>
        /// <returns>Sha1 hash.</returns>
        public static string Sha1Hash(string file)
        {
            using (FileStream fs = new FileStream(file, FileMode.Open))
            using (BufferedStream bs = new BufferedStream(fs))
            {
                using (SHA1Managed sha1 = new SHA1Managed())
                {
                    byte[] hash = sha1.ComputeHash(bs);
                    StringBuilder formatted = new StringBuilder(2 * hash.Length);
                    foreach (byte b in hash)
                    {
                        formatted.AppendFormat("{0:x2}", b);
                    }
                    return formatted.ToString();
                }
            }
        }
    }
}
