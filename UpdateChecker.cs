using System;
using System.Diagnostics;
using System.Net.Http;
using System.Threading.Tasks;

namespace PointAC
{
    public enum UpdateStatus
    {
        UpToDate,
        UpdateAvailable,
        CheckFailed
    }

    public sealed class UpdateResult
    {
        public UpdateStatus Status { get; }
        public Version? LatestVersion { get; }
        public string Message { get; }

        public bool HasUpdate => Status == UpdateStatus.UpdateAvailable;
        public bool IsFailed => Status == UpdateStatus.CheckFailed;

        public UpdateResult(UpdateStatus status, Version? latestVersion, string message)
        {
            Status = status;
            LatestVersion = latestVersion;
            Message = message;
        }
    }

    public static class UpdateChecker
    {
        /// <summary>
        /// The GitHub raw URL that contains the latest version string.
        /// Example: https://raw.githubusercontent.com/xh4l1l/Versions/refs/heads/main/PointAC
        /// </summary>
        public static string VersionUrl { get; set; } =
            "https://raw.githubusercontent.com/xh4l1l/Versions/refs/heads/main/PointAC";

        /// <summary>
        /// Checks for updates asynchronously and returns a detailed result.
        /// </summary>
        public static async Task<(UpdateStatus, Version?)> CheckForUpdatesAsync(string versionUrl, Version currentVersion)
        {
            try
            {
                using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(5) };
                string versionString = await http.GetStringAsync(versionUrl);

                if (!Version.TryParse(versionString.Trim(), out var latestVersion))
                    return (UpdateStatus.CheckFailed, null);

                if (latestVersion > currentVersion)
                    return (UpdateStatus.UpdateAvailable, latestVersion);

                return (UpdateStatus.UpToDate, latestVersion);
            }
            catch
            {
                return (UpdateStatus.CheckFailed, null);
            }
        }
    }
}