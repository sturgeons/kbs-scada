using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Globalization;
using System.Timers;

namespace DataSyncTool
{
    class Program
    {
        private static Config _config;
        private static TimeIndexManager _timeIndexManager;
        private static AccessDataReader _dataReader;
        private static AccessDataReader _ngDataReader;
        private static ParameterExtractor _parameterExtractor;
        private static MySqlDataWriter _mySqlWriter;
        private static System.Timers.Timer _timer;
        private static bool _isRunning = false;
        private static List<string> _allowedLabels = new List<string>();

        static void Main(string[] args)
        {
            Console.WriteLine("=== Access到MySQL数据同步工具 ===");
            Console.WriteLine();

            try
            {
                // 加载配置
                Console.WriteLine("正在加载配置文件...");
                _config = Config.Load("appsettings.json");
                Console.WriteLine("配置文件加载成功");

                // 初始化Time_index管理器
                _timeIndexManager = new TimeIndexManager(_config.Sync.TimeIndexFile);
                decimal lastIndex = _timeIndexManager.GetLastIndex();
                Console.WriteLine($"当前最后Time_index: {lastIndex.ToString(CultureInfo.InvariantCulture)}");

                _allowedLabels = _config.Sync.GetAllowedLabelNames();
                if (_allowedLabels.Count > 0)
                {
                    Console.WriteLine($"允许同步的Label_Name数量: {_allowedLabels.Count}");
                }
                else
                {
                    Console.WriteLine("允许同步的Label_Name未配置：将不做Label_Name过滤（全量读取）");
                }

                // 初始化Access数据读取器
                _dataReader = CreateAccessReader(_config.Access.DataPath, "合格");
                _ngDataReader = CreateAccessReader(_config.Access.NGDataPath, "不合格");

                // 测试Access连接
                Console.WriteLine("正在测试Access数据库连接...");
                if (!_dataReader.TestConnection(out string accessOkErr))
                {
                    Console.WriteLine($"警告: 无法连接到合格数据库: {_config.Access.DataPath}");
                    Console.WriteLine(accessOkErr);
                }
                else
                {
                    Console.WriteLine($"合格数据库连接成功: {_config.Access.DataPath}");
                }

                if (!_ngDataReader.TestConnection(out string accessNgErr))
                {
                    Console.WriteLine($"警告: 无法连接到不合格数据库: {_config.Access.NGDataPath}");
                    Console.WriteLine(accessNgErr);
                }
                else
                {
                    Console.WriteLine($"不合格数据库连接成功: {_config.Access.NGDataPath}");
                }

                // 初始化参数提取器
                _parameterExtractor = new ParameterExtractor(_config);

                // 初始化MySQL写入器
                string mySqlConnectionString = _config.GetMySqlConnectionString();
                Console.WriteLine($"MySQL配置: {_config.MySql.Host}:{_config.MySql.Port} / {_config.MySql.Database} (User={_config.MySql.Username})");
                _mySqlWriter = new MySqlDataWriter(mySqlConnectionString, _config);

                // 测试MySQL连接
                Console.WriteLine("正在测试MySQL数据库连接...");
                if (!_mySqlWriter.TestConnection(out string mysqlError))
                {
                    Console.WriteLine("错误: 无法连接到MySQL数据库，请检查配置");
                    Console.WriteLine("MySQL错误详情(请将这段发我):");
                    Console.WriteLine(mysqlError);
                    Console.WriteLine("按任意键退出...");
                    Console.ReadKey();
                    return;
                }
                Console.WriteLine("MySQL数据库连接成功");

                // 立即执行一次同步
                Console.WriteLine();
                Console.WriteLine("开始首次数据同步...");
                PerformSync();

                // 启动定时器
                int intervalMs = _config.Sync.IntervalMinutes * 60 * 1000;
                _timer = new System.Timers.Timer(intervalMs);
                _timer.Elapsed += Timer_Elapsed;
                _timer.AutoReset = true;
                _timer.Start();

                Console.WriteLine();
                Console.WriteLine($"定时器已启动，每 {_config.Sync.IntervalMinutes} 分钟执行一次同步");
                Console.WriteLine("按 Ctrl+C 或关闭窗口退出程序");
                Console.WriteLine();

                // 等待程序退出
                System.Threading.Thread.Sleep(System.Threading.Timeout.Infinite);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"程序启动失败: {ex.Message}");
                Console.WriteLine($"详细错误: {ex}");
                Console.WriteLine("按任意键退出...");
                Console.ReadKey();
            }
        }

        private static AccessDataReader CreateAccessReader(string dbPath, string name)
        {
            string lastError = "";
            foreach (var provider in _config.GetAccessProviderCandidates())
            {
                string cs = $"Provider={provider};Data Source={dbPath};";
                var reader = new AccessDataReader(cs);
                if (reader.TestConnection(out string err))
                {
                    Console.WriteLine($"{name}数据库使用Provider: {provider}");
                    return reader;
                }
                lastError = err;
            }

            // 都失败则返回最后一个（保持后续调用一致），并把详细错误打印出来
            Console.WriteLine($"{name}数据库未找到可用的ACE Provider。请安装 Microsoft Access Database Engine（与程序位数匹配）。");
            Console.WriteLine(lastError);
            return new AccessDataReader(_config.GetAccessConnectionString(dbPath));
        }

        private static void Timer_Elapsed(object sender, ElapsedEventArgs e)
        {
            if (_isRunning)
            {
                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 上次同步尚未完成，跳过本次执行");
                return;
            }

            PerformSync();
        }

        private static void PerformSync()
        {
            _isRunning = true;
            DateTime startTime = DateTime.Now;

            try
            {
                Console.WriteLine();
                Console.WriteLine($"[{startTime:yyyy-MM-dd HH:mm:ss}] 开始数据同步...");

                decimal lastIndex = _timeIndexManager.GetLastIndex();
                Console.WriteLine($"查询Time_index > {lastIndex.ToString(CultureInfo.InvariantCulture)} 的数据");

                // 从两个Access数据库读取数据
                List<DataRow> dataRows = new List<DataRow>();
                
                try
                {
                    var dataRows1 = _dataReader.GetDataByTimeIndex(lastIndex, _allowedLabels);
                    dataRows.AddRange(dataRows1);
                    Console.WriteLine($"从合格数据库读取到 {dataRows1.Count} 条记录");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"读取合格数据库失败: {ex.Message}");
                }

                try
                {
                    var dataRows2 = _ngDataReader.GetDataByTimeIndex(lastIndex, _allowedLabels);
                    dataRows.AddRange(dataRows2);
                    Console.WriteLine($"从不合格数据库读取到 {dataRows2.Count} 条记录");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"读取不合格数据库失败: {ex.Message}");
                }

                if (dataRows.Count == 0)
                {
                    Console.WriteLine("没有新数据需要同步");
                    return;
                }

                // 按Time_index排序
                dataRows = dataRows.OrderBy(row => ReadTimeIndex(row)).ToList();
                int totalRecords = dataRows.Count;
                int processedRecords = 0;
                Console.WriteLine($"本次待上传记录数: {totalRecords}");

                // 提取参数并插入MySQL
                decimal maxTimeIndex = lastIndex;
                int totalInserted = 0;

                foreach (DataRow row in dataRows)
                {
                    try
                    {
                        decimal currentTimeIndex = ReadTimeIndex(row);
                        if (currentTimeIndex > maxTimeIndex)
                        {
                            maxTimeIndex = currentTimeIndex;
                        }

                        // 提取参数
                        var parameters = _parameterExtractor.ExtractParameters(row);
                        
                        if (parameters.Count > 0)
                        {
                            // 插入MySQL
                            int inserted = _mySqlWriter.InsertParameters(parameters);
                            totalInserted += inserted;
                            processedRecords++;
                            int remaining = totalRecords - processedRecords;
                            Console.WriteLine($"Time_index {currentTimeIndex.ToString(CultureInfo.InvariantCulture)}: 插入 {inserted} 个参数；本次剩余未上传记录: {remaining}");
                        }
                        else
                        {
                            // 该记录未提取到参数，也算处理完成，避免剩余数不变化
                            processedRecords++;
                            int remaining = totalRecords - processedRecords;
                            Console.WriteLine($"Time_index {currentTimeIndex.ToString(CultureInfo.InvariantCulture)}: 未提取到参数；本次剩余未上传记录: {remaining}");
                        }
                    }
                    catch (Exception ex)
                    {
                        processedRecords++;
                        int remaining = totalRecords - processedRecords;
                        Console.WriteLine($"处理记录失败: {ex.Message}");
                        Console.WriteLine($"本次剩余未上传记录: {remaining}");
                    }
                }

                // 更新Time_index
                if (maxTimeIndex > lastIndex)
                {
                    _timeIndexManager.SetLastIndex(maxTimeIndex);
                    Console.WriteLine($"更新最后Time_index: {maxTimeIndex.ToString(CultureInfo.InvariantCulture)}");
                }

                DateTime endTime = DateTime.Now;
                TimeSpan duration = endTime - startTime;
                Console.WriteLine($"[{endTime:yyyy-MM-dd HH:mm:ss}] 同步完成，共处理 {dataRows.Count} 条记录，插入 {totalInserted} 个参数，耗时 {duration.TotalSeconds:F2} 秒");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 同步失败: {ex.Message}");
                Console.WriteLine($"详细错误: {ex}");
            }
            finally
            {
                _isRunning = false;
            }
        }

        private static decimal ReadTimeIndex(DataRow row)
        {
            object obj = row["Time_index"];
            if (obj == null || obj == DBNull.Value)
                return 0m;

            // Access里大概率是Double；这里用decimal保存原值，避免(long)截断导致“永远停在最后一条”
            try
            {
                if (obj is decimal dec) return dec;
                if (obj is double d) return (decimal)d;
                if (obj is float f) return (decimal)f;
                if (obj is int i) return i;
                if (obj is long l) return l;

                string s = obj.ToString()?.Trim() ?? "0";
                if (decimal.TryParse(s, NumberStyles.Number, CultureInfo.InvariantCulture, out var parsed))
                    return parsed;
                if (decimal.TryParse(s, out parsed))
                    return parsed;
            }
            catch
            {
                // ignore, fall through
            }

            return 0m;
        }
    }
}

