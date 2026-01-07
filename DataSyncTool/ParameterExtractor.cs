using System;
using System.Collections.Generic;
using System.Data;
using System.Text.RegularExpressions;

namespace DataSyncTool
{
    public class ParameterData
    {
        public string ParameterName { get; set; }
        public double Value { get; set; }
        public double LowerLimit { get; set; }
        public double UpperLimit { get; set; }
        public string PartNumber { get; set; }
        public string Variance { get; set; }
    }

    public class ParameterExtractor
    {
        private Config _config;

        public ParameterExtractor(Config config)
        {
            _config = config;
        }

        public List<ParameterData> ExtractParameters(DataRow row)
        {
            var parameters = new List<ParameterData>();

            // Label_Name白名单：不在配置清单中的直接丢弃（不写入MySQL）
            string labelName = "";
            try
            {
                if (row.Table.Columns.Contains("Label_Name") && row["Label_Name"] != DBNull.Value)
                {
                    labelName = (row["Label_Name"]?.ToString() ?? "").Trim();
                }
            }
            catch
            {
                // ignore
            }

            var allowedLabels = _config.Sync?.GetAllowedLabelNames() ?? new List<string>();
            if (allowedLabels.Count > 0)
            {
                bool ok = false;
                foreach (var allowed in allowedLabels)
                {
                    if (string.Equals(allowed, labelName, StringComparison.OrdinalIgnoreCase))
                    {
                        ok = true;
                        break;
                    }
                }

                if (!ok)
                {
                    Console.WriteLine($"Label_Name '{labelName}' 不在配置清单中，已丢弃不同步到MySQL");
                    return parameters;
                }
            }

            // 先按型号(Label_Name)映射零件号；映射不到再从Matrix_Code提取
            string partNumber = ResolvePartNumber(row);
            if (string.IsNullOrWhiteSpace(partNumber))
            {
                // 强制保证：partNumber为空的记录绝不写入MySQL
                string timeIndex = "";
                try
                {
                    if (row.Table.Columns.Contains("Time_index") && row["Time_index"] != DBNull.Value)
                        timeIndex = row["Time_index"]?.ToString() ?? "";
                }
                catch { }

                string matrixCode = "";
                try
                {
                    if (row.Table.Columns.Contains("Matrix_Code") && row["Matrix_Code"] != DBNull.Value)
                        matrixCode = row["Matrix_Code"]?.ToString() ?? "";
                }
                catch { }

                Console.WriteLine($"Time_index {timeIndex} / Label_Name '{labelName}' 的partNumber为空，已丢弃不同步到MySQL。Matrix_Code='{matrixCode}'");
                return parameters;
            }

            // 定义10个参数映射（按data.md第13-22行的顺序和名称）
            var parameterMappings = new Dictionary<string, ParameterMapping>
            {
                { "Reserve", new ParameterMapping { FieldName = "Reserve", LowerLimit = 237.7, UpperLimit = 242.3 } },
                { "3/4", new ParameterMapping { FieldName = "3_4", LowerLimit = 105.5, UpperLimit = 108.5 } },
                { "1/4", new ParameterMapping { FieldName = "1_4", LowerLimit = 213.7, UpperLimit = 218.3 } },
                { "Empty", new ParameterMapping { FieldName = "Name11Value", LowerLimit = 267.7, UpperLimit = 272.3 } },
                { "1/2", new ParameterMapping { FieldName = "1_2", LowerLimit = 159.2, UpperLimit = 162.8 } },
                { "True_Empty", new ParameterMapping { FieldName = "True_Empty", LowerLimit = 287.7, UpperLimit = 292.3 } },
                { "空载电流", new ParameterMapping { FieldName = "Current", LowerLimit = 0.3, UpperLimit = 1.9 } },
                { "泄露", new ParameterMapping { FieldName = "Supply_Leak", LowerLimit = -0.14, UpperLimit = 0.35 } },
                { "True_Full", new ParameterMapping { FieldName = "True_Full", LowerLimit = 50.5, UpperLimit = 53.5 } },
                { "Full", new ParameterMapping { FieldName = "Full", LowerLimit = 68.5, UpperLimit = 71.5 } }
            };

            foreach (var mapping in parameterMappings)
            {
                string paramName = mapping.Key;
                var paramMapping = mapping.Value;

                try
                {
                    if (row.Table.Columns.Contains(paramMapping.FieldName))
                    {
                        object valueObj = row[paramMapping.FieldName];
                        if (valueObj != null && valueObj != DBNull.Value)
                        {
                            if (double.TryParse(valueObj.ToString(), out double value))
                            {
                                string variance = _config.Variance.ContainsKey(paramName) 
                                    ? _config.Variance[paramName] 
                                    : "";

                                parameters.Add(new ParameterData
                                {
                                    ParameterName = paramName,
                                    Value = value,
                                    LowerLimit = paramMapping.LowerLimit,
                                    UpperLimit = paramMapping.UpperLimit,
                                    PartNumber = partNumber,
                                    Variance = variance
                                });
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"提取参数 {paramName} 失败: {ex.Message}");
                }
            }

            return parameters;
        }

        private string ResolvePartNumber(DataRow row)
        {
            string labelName = "";
            try
            {
                if (row.Table.Columns.Contains("Label_Name") && row["Label_Name"] != DBNull.Value)
                {
                    labelName = (row["Label_Name"]?.ToString() ?? "").Trim();
                }
            }
            catch
            {
                // ignore
            }

            // 1) 按Label_Name映射（优先）
            if (!string.IsNullOrWhiteSpace(labelName) &&
                _config.Sync?.PartNumberByLabelName != null &&
                _config.Sync.PartNumberByLabelName.TryGetValue(labelName, out var mappedPn) &&
                !string.IsNullOrWhiteSpace(mappedPn))
            {
                return mappedPn.Trim();
            }

            // 2) 从Matrix_Code提取
            string matrixCode = "";
            try
            {
                if (row.Table.Columns.Contains("Matrix_Code") && row["Matrix_Code"] != DBNull.Value)
                {
                    matrixCode = row["Matrix_Code"]?.ToString() ?? "";
                }
            }
            catch
            {
                // ignore
            }

            string extracted = ExtractPartNumber(matrixCode);
            if (!string.IsNullOrWhiteSpace(extracted))
                return extracted;

            // 3) 最后兜底：如果你用Sync.LabelName过滤了型号，则可用该型号的默认映射
            string cfgLabel = _config.Sync?.LabelName ?? "";
            if (!string.IsNullOrWhiteSpace(cfgLabel) &&
                _config.Sync?.PartNumberByLabelName != null &&
                _config.Sync.PartNumberByLabelName.TryGetValue(cfgLabel, out var cfgPn) &&
                !string.IsNullOrWhiteSpace(cfgPn))
            {
                return cfgPn.Trim();
            }

            return "";
        }

        private string ExtractPartNumber(string matrixCode)
        {
            if (string.IsNullOrEmpty(matrixCode))
                return "";

            // 使用正则表达式提取partNumber，例如从 "#5QD919051T    ###*539 M2G1JJP8252*=" 中提取 "5QD919051T"
            var match = Regex.Match(matrixCode, @"#\s*([A-Z0-9]+)");
            if (match.Success && match.Groups.Count > 1)
            {
                return match.Groups[1].Value;
            }

            return "";
        }

        private class ParameterMapping
        {
            public string FieldName { get; set; }
            public double LowerLimit { get; set; }
            public double UpperLimit { get; set; }
        }
    }
}

