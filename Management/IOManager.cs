using PointAC.Core;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows;

namespace PointAC.Management
{
    public static class IOManager
    {
        public static string FileType { get; } = ".pac";

        public class AppFileData
        {
            public int Duration { get; set; }
            public List<PointEntry> Points { get; set; } = new();
        }

        public static async Task<AppFileData> LoadFromFileAsync(string filePath)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(filePath))
                    throw new ArgumentException("File path is invalid.");

                if (!File.Exists(filePath))
                    throw new FileNotFoundException("PAC file not found.", filePath);

                string json = await File.ReadAllTextAsync(filePath);
                if (string.IsNullOrWhiteSpace(json))
                    throw new InvalidDataException("PAC file is empty.");

                var data = JsonSerializer.Deserialize<AppFileData>(json, new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true
                });

                if (data == null)
                    throw new InvalidOperationException("Invalid PAC file format.");

                if (data.Points == null || data.Points.Count == 0)
                    throw new InvalidDataException("PAC file contains no valid points.");

                return data;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"[LoadFromFileAsync]\nFilePath: {filePath}\n\n{ex}", "DEBUG ERROR", MessageBoxButton.OK, MessageBoxImage.Error);
                throw;
            }
        }

        public static async Task SaveToFileAsync(string filePath, AppFileData data)
        {
            if (data == null)
                throw new ArgumentNullException(nameof(data));

            bool indent = false;

            #if DEBUG
            indent = true;
            #endif

            string json = JsonSerializer.Serialize(data, new JsonSerializerOptions
            {
                WriteIndented = indent
            });

            await File.WriteAllTextAsync(filePath, json);
        }
    }
}