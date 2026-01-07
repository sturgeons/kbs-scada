using System;
using System.Data;
using System.Data.OleDb;
using System.Collections.Generic;
using System.Linq;

namespace DataSyncTool
{
    public class AccessDataReader
    {
        private string _connectionString;

        public AccessDataReader(string connectionString)
        {
            _connectionString = connectionString;
        }

        public List<DataRow> GetDataByTimeIndex(decimal lastTimeIndex, string labelName = "")
        {
            if (string.IsNullOrWhiteSpace(labelName))
                return GetDataByTimeIndex(lastTimeIndex, Enumerable.Empty<string>());

            return GetDataByTimeIndex(lastTimeIndex, new[] { labelName });
        }

        public List<DataRow> GetDataByTimeIndex(decimal lastTimeIndex, IEnumerable<string> labelNames)
        {
            var results = new List<DataRow>();

            try
            {
                using (var connection = new OleDbConnection(_connectionString))
                {
                    connection.Open();

                    // 用参数化，避免Time_index为Double/Decimal时的精度与类型问题
                    string query = "SELECT * FROM Data WHERE Time_index > ?";

                    var labels = (labelNames ?? Enumerable.Empty<string>())
                        .Where(x => !string.IsNullOrWhiteSpace(x))
                        .Select(x => x.Trim())
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .ToList();

                    if (labels.Count > 0)
                    {
                        // OleDb不支持命名参数，使用位置参数 ?，这里拼 OR
                        query += " AND (" + string.Join(" OR ", labels.Select(_ => "Label_Name = ?")) + ")";
                    }
                    query += " ORDER BY Time_index";

                    using (var command = new OleDbCommand(query, connection))
                    {
                        // Access大概率是Double（Time_index>2^31），用Double参数最稳
                        command.Parameters.Add(new OleDbParameter
                        {
                            OleDbType = OleDbType.Double,
                            Value = (double)lastTimeIndex
                        });
                        if (labels.Count > 0)
                        {
                            foreach (var label in labels)
                            {
                                command.Parameters.Add(new OleDbParameter
                                {
                                    OleDbType = OleDbType.VarWChar,
                                    Value = label
                                });
                            }
                        }

                        using (var adapter = new OleDbDataAdapter(command))
                        {
                            var dataTable = new DataTable();
                            adapter.Fill(dataTable);

                            foreach (DataRow row in dataTable.Rows)
                            {
                                results.Add(row);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"读取Access数据库失败: {ex.Message}", ex);
            }

            return results;
        }

        public bool TestConnection()
        {
            try
            {
                using (var connection = new OleDbConnection(_connectionString))
                {
                    connection.Open();
                    return true;
                }
            }
            catch
            {
                return false;
            }
        }

        public bool TestConnection(out string error)
        {
            try
            {
                using (var connection = new OleDbConnection(_connectionString))
                {
                    connection.Open();
                    error = "";
                    return true;
                }
            }
            catch (Exception ex)
            {
                error = ex.ToString();
                return false;
            }
        }
    }
}

