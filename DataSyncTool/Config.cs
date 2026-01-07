using MySql.Data.MySqlClient;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;

namespace DataSyncTool
{
    public class Config
    {
        public MySqlConfig MySql { get; set; }
        public AccessConfig Access { get; set; }
        public SyncConfig Sync { get; set; }
        public FixedFieldsConfig FixedFields { get; set; }
        public Dictionary<string, string> Variance { get; set; }

        public static Config Load(string configPath = "appsettings.json")
        {
            try
            {
                if (!File.Exists(configPath))
                {
                    throw new FileNotFoundException($"配置文件不存在: {configPath}");
                }

                string json = File.ReadAllText(configPath);
                var config = JsonConvert.DeserializeObject<Config>(json);
                
                if (config == null)
                {
                    throw new Exception("配置文件解析失败");
                }

                return config;
            }
            catch (Exception ex)
            {
                throw new Exception($"加载配置文件失败: {ex.Message}", ex);
            }
        }

        public string GetMySqlConnectionString()
        {
            // 使用Builder可正确处理密码/用户名中的特殊字符（如 ; = 等）
            var builder = new MySqlConnectionStringBuilder
            {
                Server = MySql.Host,
                Port = (uint)MySql.Port,
                Database = MySql.Database,
                UserID = MySql.Username,
                Password = MySql.Password,
                CharacterSet = "utf8mb4",
                // MySQL8常见：caching_sha2_password在非SSL场景可能需要AllowPublicKeyRetrieval
                AllowPublicKeyRetrieval = MySql.AllowPublicKeyRetrieval,
                SslMode = ParseSslMode(MySql.SslMode),
                ConnectionTimeout = (uint)Math.Max(1, MySql.ConnectionTimeoutSeconds),
                DefaultCommandTimeout = (uint)Math.Max(1, MySql.DefaultCommandTimeoutSeconds),
                Pooling = true
            };

            return builder.ConnectionString;
        }

        public string GetAccessConnectionString(string dbPath)
        {
            // 默认优先用16.0（Access Database Engine 2016），如果用户指定则用指定值
            string provider = (Access?.Provider ?? "").Trim();
            if (string.IsNullOrWhiteSpace(provider))
                provider = "Microsoft.ACE.OLEDB.16.0";

            return $"Provider={provider};Data Source={dbPath};";
        }

        public List<string> GetAccessProviderCandidates()
        {
            string provider = (Access?.Provider ?? "").Trim();
            if (!string.IsNullOrWhiteSpace(provider))
                return new List<string> { provider };

            return new List<string>
            {
                "Microsoft.ACE.OLEDB.16.0",
                "Microsoft.ACE.OLEDB.12.0"
            };
        }

        private static MySqlSslMode ParseSslMode(string sslMode)
        {
            // 兼容配置里写 "None/Preferred/Required/VerifyCA/VerifyFull"（不区分大小写）
            if (string.IsNullOrWhiteSpace(sslMode))
                return MySqlSslMode.Preferred;

            if (Enum.TryParse<MySqlSslMode>(sslMode, ignoreCase: true, out var mode))
                return mode;

            return MySqlSslMode.Preferred;
        }
    }

    public class MySqlConfig
    {
        public string Host { get; set; }
        public int Port { get; set; }
        public string Database { get; set; }
        public string Username { get; set; }
        public string Password { get; set; }

        // 可选：MySQL8/SSL相关参数（不写则走默认）
        public string SslMode { get; set; } = "Preferred";
        public bool AllowPublicKeyRetrieval { get; set; } = true;
        public int ConnectionTimeoutSeconds { get; set; } = 5;
        public int DefaultCommandTimeoutSeconds { get; set; } = 30;
    }

    public class AccessConfig
    {
        public string DataPath { get; set; }
        public string NGDataPath { get; set; }

        // 可选：指定Access Provider（如 "Microsoft.ACE.OLEDB.16.0" 或 "Microsoft.ACE.OLEDB.12.0"）
        // 为空则自动尝试 16.0 -> 12.0
        public string Provider { get; set; } = "";
    }

    public class SyncConfig
    {
        public int IntervalMinutes { get; set; }
        public string TimeIndexFile { get; set; }

        // 可选：只同步指定Label_Name（为空则不过滤）
        // 之前代码里写死为 'A1C6/BC316 T'，这里保留默认值，避免升级后行为变化
        public string LabelName { get; set; } = "A1C6/BC316 T";

        // 可选：按型号(Label_Name)指定零件号(partNumber)
        // 例如：当Label_Name为"A1C6/BC316 T"时，partNumber固定写入"5QD919051T"
        public Dictionary<string, string> PartNumberByLabelName { get; set; } =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "A1C6/BC316 T", "5QD919051T" }
            };

        /// <summary>
        /// 返回允许同步的Label_Name清单。
        /// 规则：优先使用 PartNumberByLabelName 的键；如果为空则使用 LabelName（单值）。
        /// 若最终为空，则表示不做Label_Name限制（全量）。
        /// </summary>
        public List<string> GetAllowedLabelNames()
        {
            var list = new List<string>();

            if (PartNumberByLabelName != null && PartNumberByLabelName.Count > 0)
            {
                foreach (var key in PartNumberByLabelName.Keys)
                {
                    if (!string.IsNullOrWhiteSpace(key))
                        list.Add(key.Trim());
                }
            }
            else if (!string.IsNullOrWhiteSpace(LabelName))
            {
                list.Add(LabelName.Trim());
            }

            // 去重（不区分大小写）
            var dedup = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var result = new List<string>();
            foreach (var s in list)
            {
                if (dedup.Add(s))
                    result.Add(s);
            }
            return result;
        }
    }

    public class FixedFieldsConfig
    {
        public string Supplier { get; set; }
        public string SupplierName { get; set; }
        public string VehicleType { get; set; }
        public string PartName { get; set; }
        public string Station { get; set; }
        public string BTV { get; set; }
    }
}

