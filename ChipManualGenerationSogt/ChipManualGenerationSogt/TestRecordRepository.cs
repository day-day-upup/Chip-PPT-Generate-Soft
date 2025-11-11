using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient; // ? 改为 .NET Framework 兼容命名空间
using System.Linq;
using System.Text;

namespace ChipManualGenerationSogt
{
    public class TestRecord
    {
        public int? ID { get; set; }
        public string PN { get; set; }
        public string ON1 { get; set; }
        public string SN { get; set; }
        public string Temp { get; set; }
        public string Tester { get; set; }
        public DateTime? TestTime { get; set; }
        public string TestResult { get; set; }
        public string StatusC { get; set; }
        public string E2Record { get; set; }

        // InVSWR
        public string InVSWR1 { get; set; }
        public string InVSWR2 { get; set; }
        public string InVSWR3 { get; set; }
        public string InVSWR4 { get; set; }

        // OutVSWR
        public string OutVSWR1 { get; set; }
        public string OutVSWR2 { get; set; }
        public string OutVSWR3 { get; set; }
        public string OutVSWR4 { get; set; }

        // Gain
        public string Gain1 { get; set; }
        public string Gain2 { get; set; }
        public string Gain3 { get; set; }
        public string Gain4 { get; set; }

        // GainMax
        public string GainMax1 { get; set; }
        public string GainMax2 { get; set; }
        public string GainMax3 { get; set; }
        public string GainMax4 { get; set; }

        // GainIm
        public string GainIm1 { get; set; }
        public string GainIm2 { get; set; }
        public string GainIm3 { get; set; }
        public string GainIm4 { get; set; }

        // Isolation
        public string Isolation1 { get; set; }
        public string Isolation2 { get; set; }
        public string Isolation3 { get; set; }
        public string Isolation4 { get; set; }

        // NF
        public string NF1 { get; set; }
        public string NF2 { get; set; }
        public string NF3 { get; set; }
        public string NF4 { get; set; }

        // OIP3
        public string OIP31 { get; set; }
        public string OIP32 { get; set; }
        public string OIP33 { get; set; }
        public string OIP34 { get; set; }

        // LIM3
        public string LIM31 { get; set; }
        public string LIM32 { get; set; }
        public string LIM33 { get; set; }
        public string LIM34 { get; set; }

        // RIM3
        public string RIM31 { get; set; }
        public string RIM32 { get; set; }
        public string RIM33 { get; set; }
        public string RIM34 { get; set; }

        // P1dB
        public string P1dB1 { get; set; }
        public string P1dB2 { get; set; }
        public string P1dB3 { get; set; }
        public string P1dB4 { get; set; }

        // Psat
        public string Psat1 { get; set; }
        public string Psat2 { get; set; }
        public string Psat3 { get; set; }
        public string Psat4 { get; set; }

        // Current
        public string Current1 { get; set; }
        public string Current2 { get; set; }
        public string Current3 { get; set; }
        public string Current4 { get; set; }

        public string SoftVersion { get; set; }
    }

    public class TaskModel_DB
    {
        public int ID { get; set; }           // 对应 int
        public string TaskName { get; set; }  // 对应 nvarchar(MAX)
        public int Status { get; set; }       // 对应 int
        public DateTime StartDate { get; set; } // 对应 datetime
        public int Consumed { get; set; }     // 对应 int
        public bool DataStatus { get; set; }  // 对应 bit
        public bool FilesStatus { get; set; } // 对应 bit
        public bool CheckStatus { get; set; } // 对应 bit
    }

    public class User
    { 
        public int ID { get; set; }
        public string UserName { get; set; }

        public string Password { get; set; }

        public int priority { get; set; } //0 - 管理员，1 - 提供数据，2 -制造ppt，3 - 审核数据
    }
    public class TestRecordRepository
    {
        private string _connectionString;

        public TestRecordRepository(string connectionString)
        {
            if (string.IsNullOrEmpty(connectionString))
                throw new ArgumentNullException("connectionString");

            _connectionString = connectionString;
        }

        

        /// <summary>
        /// 根据 PN 列表和可选时间范围查询测试记录（最多1000条）
        /// </summary>
        public List<TestRecord> GetRecordsByPN(
            string[] keywords,
            DateTime? startTime = null,
            DateTime? endTime = null)
        {
            if (keywords == null || keywords.Length == 0)
                throw new ArgumentException("关键词列表不能为空", "keywords");

            StringBuilder sql = new StringBuilder();
            List<SqlParameter> parameters = new List<SqlParameter>();

            // 构建多个 LIKE 条件
            string[] likeConditions = new string[keywords.Length];
            for (int i = 0; i < keywords.Length; i++)
            {
                string paramName = "@Keyword" + i;
                likeConditions[i] = "[PN] LIKE " + paramName;
                parameters.Add(new SqlParameter(paramName, "%" + keywords[i] + "%"));
            }

            sql.Append(@"
                SELECT TOP (1000)
                    [ID], [PN], [ON1], [SN], [Temp], [Tester], [TestTime], [TestResult], [StatusC], [E2Record],
                    [InVSWR1], [InVSWR2], [InVSWR3], [InVSWR4],
                    [OutVSWR1], [OutVSWR2], [OutVSWR3], [OutVSWR4],
                    [Gain1], [Gain2], [Gain3], [Gain4],
                    [GainMax1], [GainMax2], [GainMax3], [GainMax4],
                    [GainIm1], [GainIm2], [GainIm3], [GainIm4],
                    [Isolation1], [Isolation2], [Isolation3], [Isolation4],
                    [NF1], [NF2], [NF3], [NF4],
                    [OIP31], [OIP32], [OIP33], [OIP34],
                    [LIM31], [LIM32], [LIM33], [LIM34],
                    [RIM31], [RIM32], [RIM33], [RIM34],
                    [P1dB1], [P1dB2], [P1dB3], [P1dB4],
                    [Psat1], [Psat2], [Psat3], [Psat4],
                    [Current1], [Current2], [Current3], [Current4],
                    [SoftVersion]
                FROM [QotanaTestSystem].[dbo].[T31_TestRecord]
                WHERE (");

            sql.Append(string.Join(" OR ", likeConditions));
            sql.Append(")");

            if (startTime.HasValue)
            {
                sql.Append(" AND [TestTime] >= @StartTime");
                parameters.Add(new SqlParameter("@StartTime", startTime.Value));
            }

            if (endTime.HasValue)
            {
                sql.Append(" AND [TestTime] < @EndTime");
                parameters.Add(new SqlParameter("@EndTime", endTime.Value));
            }

            sql.Append(" ORDER BY [TestTime] DESC");

            return ExecuteQuery(sql.ToString(), parameters.ToArray());
        }

        public List<User> GetUsers(string connectionString)
        { 
                var users = new List<User>();
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    using (var command = new SqlCommand())
                    {
                        command.Connection = connection;
                        command.CommandType = CommandType.Text;
                        command.CommandText = @"  
                                SELECT  *
                                FROM yp_test_user; 
		                         ";

                        SqlDataReader reader = command.ExecuteReader();

                        while (reader.Read())
                        {
                            var user = new User();
                            // 安全读取可空字段（PersonID、StoreID、TerritoryID 可能为 NULL）
                            user.ID = Convert.ToInt32( reader["ID"]  );
                            user.UserName = reader["UserName"] as string;
                            user.Password = reader["Password"] as string;
                            user.priority = Convert.ToInt32(reader["priority"] );
                            users.Add(user);
                        }
                    }

                    
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
                return users;
            }
        
        }

        private List<TestRecord> ExecuteQuery(string sql, SqlParameter[] parameters)
        {
            List<TestRecord> records = new List<TestRecord>();

            using (SqlConnection connection = new SqlConnection(_connectionString))
            {
                try
                {
                    connection.Open();
                    Console.WriteLine("Connected successfully.");

                    using (SqlCommand command = new SqlCommand(sql, connection))
                    {
                        command.Parameters.AddRange(parameters);

                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                int? id = null;
                                object idValue = reader["ID"];
                                if (idValue != DBNull.Value)
                                {
                                    int parsedId;
                                    if (int.TryParse(idValue.ToString(), out parsedId))
                                        id = parsedId;
                                }

                                TestRecord record = new TestRecord();
                                record.ID = id;
                                record.PN = reader["PN"] as string;
                                record.ON1 = reader["ON1"] as string;
                                record.SN = reader["SN"] as string;
                                record.Temp = reader["Temp"] as string;
                                record.Tester = reader["Tester"] as string;
                                record.TestTime = reader["TestTime"] == DBNull.Value ? (DateTime?)null : (DateTime)reader["TestTime"];
                                record.TestResult = reader["TestResult"] as string;
                                record.StatusC = reader["StatusC"] as string;
                                record.E2Record = reader["E2Record"] as string;

                                record.InVSWR1 = reader["InVSWR1"] as string;
                                record.InVSWR2 = reader["InVSWR2"] as string;
                                record.InVSWR3 = reader["InVSWR3"] as string;
                                record.InVSWR4 = reader["InVSWR4"] as string;

                                record.OutVSWR1 = reader["OutVSWR1"] as string;
                                record.OutVSWR2 = reader["OutVSWR2"] as string;
                                record.OutVSWR3 = reader["OutVSWR3"] as string;
                                record.OutVSWR4 = reader["OutVSWR4"] as string;

                                record.Gain1 = reader["Gain1"] as string;
                                record.Gain2 = reader["Gain2"] as string;
                                record.Gain3 = reader["Gain3"] as string;
                                record.Gain4 = reader["Gain4"] as string;

                                record.GainMax1 = reader["GainMax1"] as string;
                                record.GainMax2 = reader["GainMax2"] as string;
                                record.GainMax3 = reader["GainMax3"] as string;
                                record.GainMax4 = reader["GainMax4"] as string;

                                record.GainIm1 = reader["GainIm1"] as string;
                                record.GainIm2 = reader["GainIm2"] as string;
                                record.GainIm3 = reader["GainIm3"] as string;
                                record.GainIm4 = reader["GainIm4"] as string;

                                record.Isolation1 = reader["Isolation1"] as string;
                                record.Isolation2 = reader["Isolation2"] as string;
                                record.Isolation3 = reader["Isolation3"] as string;
                                record.Isolation4 = reader["Isolation4"] as string;

                                record.NF1 = reader["NF1"] as string;
                                record.NF2 = reader["NF2"] as string;
                                record.NF3 = reader["NF3"] as string;
                                record.NF4 = reader["NF4"] as string;

                                record.OIP31 = reader["OIP31"] as string;
                                record.OIP32 = reader["OIP32"] as string;
                                record.OIP33 = reader["OIP33"] as string;
                                record.OIP34 = reader["OIP34"] as string;

                                record.LIM31 = reader["LIM31"] as string;
                                record.LIM32 = reader["LIM32"] as string;
                                record.LIM33 = reader["LIM33"] as string;
                                record.LIM34 = reader["LIM34"] as string;

                                record.RIM31 = reader["RIM31"] as string;
                                record.RIM32 = reader["RIM32"] as string;
                                record.RIM33 = reader["RIM33"] as string;
                                record.RIM34 = reader["RIM34"] as string;

                                record.P1dB1 = reader["P1dB1"] as string;
                                record.P1dB2 = reader["P1dB2"] as string;
                                record.P1dB3 = reader["P1dB3"] as string;
                                record.P1dB4 = reader["P1dB4"] as string;

                                record.Psat1 = reader["Psat1"] as string;
                                record.Psat2 = reader["Psat2"] as string;
                                record.Psat3 = reader["Psat3"] as string;
                                record.Psat4 = reader["Psat4"] as string;

                                record.Current1 = reader["Current1"] as string;
                                record.Current2 = reader["Current2"] as string;
                                record.Current3 = reader["Current3"] as string;
                                record.Current4 = reader["Current4"] as string;

                                record.SoftVersion = reader["SoftVersion"] as string;

                                records.Add(record);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error connecting to database: " + ex.Message);
                }
            }

            return records;
        }


        // 写入yp_test_finished_task 表
        public void SaveTask(string connectionString, TaskModel_DB task)
        {
            const string sql = @"
        INSERT INTO yp_test_finished_task (ID, TaskName, Status, StartDate, Consumed, DataStatus, FilesStatus, CheckStatus)
        VALUES (@ID, @TaskName, @Status, @StartDate, @Consumed, @DataStatus, @FilesStatus, @CheckStatus)";

             var connection = new SqlConnection(connectionString);
             var command = new SqlCommand(sql, connection);

            // 参数化赋值（注意：所有字段都不允许 NULL，必须提供有效值）
            command.Parameters.AddWithValue("@ID", task.ID);
            command.Parameters.AddWithValue("@TaskName", task.TaskName ?? string.Empty); // 防止 null
            command.Parameters.AddWithValue("@Status", task.Status);
            command.Parameters.AddWithValue("@StartDate", task.StartDate);
            command.Parameters.AddWithValue("@Consumed", task.Consumed);
            command.Parameters.AddWithValue("@DataStatus", task.DataStatus);
            command.Parameters.AddWithValue("@FilesStatus", task.FilesStatus);
            command.Parameters.AddWithValue("@CheckStatus", task.CheckStatus);

            try
            {
                connection.Open();
                int rowsAffected = command.ExecuteNonQuery();
                Console.WriteLine($"成功插入 {rowsAffected} 行。");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"保存任务失败: {ex.Message}");
                throw;
            }
        }

        //写入yp_test_current_task 表

        public void SaveTask_CurrentTask(string connectionString, TaskModel_DB task)
        {
            const string sql = @"
        INSERT INTO yp_test_current_task (ID, TaskName, Status, StartDate, Consumed, DataStatus, FilesStatus, CheckStatus)
        VALUES (@ID, @TaskName, @Status, @StartDate, @Consumed, @DataStatus, @FilesStatus, @CheckStatus)";

            var connection = new SqlConnection(connectionString);
            var command = new SqlCommand(sql, connection);

            // 参数化赋值（注意：所有字段都不允许 NULL，必须提供有效值）
            command.Parameters.AddWithValue("@ID", task.ID);
            command.Parameters.AddWithValue("@TaskName", task.TaskName ?? string.Empty); // 防止 null
            command.Parameters.AddWithValue("@Status", task.Status);
            command.Parameters.AddWithValue("@StartDate", task.StartDate);
            command.Parameters.AddWithValue("@Consumed", task.Consumed);
            command.Parameters.AddWithValue("@DataStatus", task.DataStatus);
            command.Parameters.AddWithValue("@FilesStatus", task.FilesStatus);
            command.Parameters.AddWithValue("@CheckStatus", task.CheckStatus);

            try
            {
                connection.Open();
                int rowsAffected = command.ExecuteNonQuery();
                Console.WriteLine($"成功插入 {rowsAffected} 行。");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"保存任务失败: {ex.Message}");
                throw;
            }
        }

       
    }
}