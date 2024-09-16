using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.Linq;
using System.Threading.Tasks;

namespace MessageIngest
{
    public static class StringSimilarity
{
    public static double CalculateSimilarity(string source, string target)
    {
        if (string.IsNullOrEmpty(source) && string.IsNullOrEmpty(target))
            return 1.0; // Both strings are null or empty, consider as an exact match

        if (string.IsNullOrEmpty(source) || string.IsNullOrEmpty(target))
            return 0.0; // One of the strings is null or empty, no similarity

        int distance = LevenshteinDistance(source, target);
        int maxLen = Math.Max(source.Length, target.Length);

        // Normalize to a value between 0 and 1
        double result = 1.0 - (double)distance / maxLen;
    //    Console.WriteLine("Singularity s:" + source + " [" + result + "] " + target );
        return result;
    }

    private static int LevenshteinDistance(string source, string target)
    {
        int[,] dp = new int[source.Length + 1, target.Length + 1];

        for (int i = 0; i <= source.Length; i++)
            dp[i, 0] = i;
        for (int j = 0; j <= target.Length; j++)
            dp[0, j] = j;

        for (int i = 1; i <= source.Length; i++)
        {
            for (int j = 1; j <= target.Length; j++)
            {
                int cost = (source[i - 1] == target[j - 1]) ? 0 : 1;
                dp[i, j] = Math.Min(
                    Math.Min(dp[i - 1, j] + 1, dp[i, j - 1] + 1),
                    dp[i - 1, j - 1] + cost);
            }
        }

        return dp[source.Length, target.Length];
    }
}

}