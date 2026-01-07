using System;
using System.Globalization;
using System.IO;

namespace DataSyncTool
{
    public class TimeIndexManager
    {
        private string _filePath;
        private decimal _lastIndex;

        public TimeIndexManager(string filePath)
        {
            _filePath = filePath;
            LoadLastIndex();
        }

        public decimal GetLastIndex()
        {
            return _lastIndex;
        }

        public void SetLastIndex(decimal index)
        {
            _lastIndex = index;
            SaveLastIndex();
        }

        private void LoadLastIndex()
        {
            try
            {
                if (File.Exists(_filePath))
                {
                    string content = File.ReadAllText(_filePath).Trim();
                    if (decimal.TryParse(content, NumberStyles.Number, CultureInfo.InvariantCulture, out decimal index))
                    {
                        _lastIndex = index;
                    }
                    else
                    {
                        _lastIndex = 0m;
                    }
                }
                else
                {
                    _lastIndex = 0m;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"加载Time_index失败: {ex.Message}");
                _lastIndex = 0m;
            }
        }

        private void SaveLastIndex()
        {
            try
            {
                File.WriteAllText(_filePath, _lastIndex.ToString(CultureInfo.InvariantCulture));
            }
            catch (Exception ex)
            {
                Console.WriteLine($"保存Time_index失败: {ex.Message}");
            }
        }
    }
}

