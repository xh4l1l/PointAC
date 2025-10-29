using System.Net.Http;
using System.Diagnostics;

namespace PointAC.Services
{
    public static class UpdateService
    {

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
    }
}