using System;
using System.Net.Http;
using System.Reflection;
using System.Text.Json;
using System.Threading.Tasks;

namespace CopyForLLM
{
    public class UpdateCheckResult
    {
        public bool Success { get; set; }
        public bool UpdateAvailable { get; set; }
        public string CurrentVersion { get; set; }
        public string LatestVersion { get; set; }
        public string ReleaseUrl { get; set; }
        public string ErrorMessage { get; set; }
    }

    public static class VersionChecker
    {
        private const string GitHubApiUrl = "https://api.github.com/repos/dylanschweitzer/excel-addin/releases/latest";
        private static readonly TimeSpan Timeout = TimeSpan.FromSeconds(10);

        public static string GetCurrentVersion()
        {
            return Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? "1.2.0.0";
        }

        public static async Task<UpdateCheckResult> CheckForUpdatesAsync()
        {
            var result = new UpdateCheckResult
            {
                CurrentVersion = GetCurrentVersion()
            };

            try
            {
                using (var client = new HttpClient())
                {
                    client.Timeout = Timeout;
                    client.DefaultRequestHeaders.Add("User-Agent", "CopyForLLM-UpdateChecker");
                    client.DefaultRequestHeaders.Add("Accept", "application/vnd.github.v3+json");

                    var response = await client.GetAsync(GitHubApiUrl);
                    response.EnsureSuccessStatusCode();

                    var json = await response.Content.ReadAsStringAsync();
                    using var doc = JsonDocument.Parse(json);
                    var root = doc.RootElement;

                    if (root.TryGetProperty("tag_name", out var tagName))
                    {
                        result.LatestVersion = tagName.GetString();
                    }

                    if (root.TryGetProperty("html_url", out var htmlUrl))
                    {
                        result.ReleaseUrl = htmlUrl.GetString();
                    }

                    result.Success = true;
                    result.UpdateAvailable = IsNewerVersion(result.CurrentVersion, result.LatestVersion);
                }
            }
            catch (TaskCanceledException)
            {
                result.Success = false;
                result.ErrorMessage = "The request timed out. Please check your internet connection.";
            }
            catch (HttpRequestException ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Network error: {ex.Message}";
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error checking for updates: {ex.Message}";
            }

            return result;
        }

        private static Version NormalizeVersion(string versionString)
        {
            if (string.IsNullOrEmpty(versionString))
                return new Version(0, 0, 0, 0);

            // Remove leading 'v' if present (e.g., "v1.0.1" -> "1.0.1")
            var normalized = versionString.TrimStart('v', 'V');

            if (Version.TryParse(normalized, out var version))
            {
                return version;
            }

            return new Version(0, 0, 0, 0);
        }

        private static bool IsNewerVersion(string currentVersionString, string latestVersionString)
        {
            var currentVersion = NormalizeVersion(currentVersionString);
            var latestVersion = NormalizeVersion(latestVersionString);

            return latestVersion > currentVersion;
        }
    }
}
