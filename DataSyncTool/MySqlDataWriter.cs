using System;
using System.Collections.Generic;
using MySql.Data.MySqlClient;

namespace DataSyncTool
{
    public class MySqlDataWriter
    {
        private string _connectionString;
        private Config _config;

        public MySqlDataWriter(string connectionString, Config config)
        {
            _connectionString = connectionString;
            _config = config;
        }

        public int InsertParameters(List<ParameterData> parameters)
        {
            if (parameters == null || parameters.Count == 0)
                return 0;

            int insertedCount = 0;

            try
            {
                using (var connection = new MySqlConnection(_connectionString))
                {
                    connection.Open();
                    using (var transaction = connection.BeginTransaction())
                    {
                        try
                        {
                            string sql = @"INSERT INTO vw.parameters 
                                (supplier, supplierName, vehicleType, partNumber, partName, variance, station, `parameter`, value, lower_limit, upper_limit, inTime, BTV)
                                VALUES 
                                (@supplier, @supplierName, @vehicleType, @partNumber, @partName, @variance, @station, @parameter, @value, @lower_limit, @upper_limit, NOW(), @BTV)";

                            using (var command = new MySqlCommand(sql, connection, transaction))
                            {
                                foreach (var param in parameters)
                                {
                                    // 保险：partNumber为空的一律不写入
                                    if (string.IsNullOrWhiteSpace(param.PartNumber))
                                    {
                                        Console.WriteLine($"跳过写入：partNumber为空，parameter={param.ParameterName}");
                                        continue;
                                    }

                                    command.Parameters.Clear();
                                    command.Parameters.AddWithValue("@supplier", _config.FixedFields.Supplier);
                                    command.Parameters.AddWithValue("@supplierName", _config.FixedFields.SupplierName);
                                    command.Parameters.AddWithValue("@vehicleType", _config.FixedFields.VehicleType);
                                    command.Parameters.AddWithValue("@partNumber", param.PartNumber.Trim());
                                    command.Parameters.AddWithValue("@partName", _config.FixedFields.PartName);
                                    command.Parameters.AddWithValue("@variance", param.Variance);
                                    command.Parameters.AddWithValue("@station", _config.FixedFields.Station);
                                    command.Parameters.AddWithValue("@parameter", param.ParameterName);
                                    command.Parameters.AddWithValue("@value", param.Value);
                                    command.Parameters.AddWithValue("@lower_limit", param.LowerLimit);
                                    command.Parameters.AddWithValue("@upper_limit", param.UpperLimit);
                                    command.Parameters.AddWithValue("@BTV", _config.FixedFields.BTV);

                                    command.ExecuteNonQuery();
                                    insertedCount++;
                                }
                            }

                            transaction.Commit();
                        }
                        catch
                        {
                            transaction.Rollback();
                            throw;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"插入MySQL数据失败: {ex.Message}", ex);
            }

            return insertedCount;
        }

        public bool TestConnection()
        {
            try
            {
                using (var connection = new MySqlConnection(_connectionString))
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
                using (var connection = new MySqlConnection(_connectionString))
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

