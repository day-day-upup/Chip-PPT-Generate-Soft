
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using QC = Microsoft.Data.SqlClient; // 使用别名

namespace ConsoleApp1
{
    public class TestRecord
    {
        public int? ID { get; set; }
        public string PN { get; set; }
        public string ON1 { get; set; }
        public string SN { get; set; }
        public string Temp { get; set; } // 改为可空，更安全
        public string Tester { get; set; }
        public DateTime? TestTime { get; set; } // 建议可空
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

        // NF (Noise Figure)
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

    public class TestRecordRepository
    {
        private readonly string _connectionString;

        public TestRecordRepository(string connectionString)
        {
            _connectionString = connectionString ?? throw new ArgumentNullException(nameof(connectionString));
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
                throw new ArgumentException("关键词列表不能为空", nameof(keywords));

            var sql = new StringBuilder();
            var parameters = new List<QC.SqlParameter>();

            // 构建多个 LIKE 条件
            var likeConditions = new string[keywords.Length];
            for (int i = 0; i < keywords.Length; i++)
            {
                string paramName = $"@Keyword{i}";
                likeConditions[i] = $"[PN] LIKE {paramName}";
                parameters.Add(new QC.SqlParameter(paramName, $"%{keywords[i]}%"));
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
                parameters.Add(new QC.SqlParameter("@StartTime", startTime.Value));
            }

            if (endTime.HasValue)
            {
                sql.Append(" AND [TestTime] < @EndTime");
                parameters.Add(new QC.SqlParameter("@EndTime", endTime.Value));
            }

            sql.Append(" ORDER BY [TestTime] DESC");

            return ExecuteQuery(sql.ToString(), parameters.ToArray());
        }

        private List<TestRecord> ExecuteQuery(string sql, QC.SqlParameter[] parameters)
        {
            var records = new List<TestRecord>();

            using (var connection = new QC.SqlConnection(_connectionString))
            {
                try
                {
                    connection.Open();
                    Console.WriteLine("Connected successfully.");

                    using (var command = new QC.SqlCommand(sql, connection))
                    {
                        command.Parameters.AddRange(parameters);

                        using (var reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                records.Add(new TestRecord
                                {
                                    ID = reader.IsDBNull("ID") ? null : (int?)reader.GetInt32("ID"),
                                    PN = reader.IsDBNull("PN") ? null : reader.GetString("PN"),
                                    ON1 = reader.IsDBNull("ON1") ? null : reader.GetString("ON1"),
                                    SN = reader.IsDBNull("SN") ? null : reader.GetString("Temp"),
                                    Temp = reader.IsDBNull("Temp") ? null : reader.GetString("SN"),
                                    Tester = reader.IsDBNull("Tester") ? null : reader.GetString("Tester"),
                                    TestTime = reader.IsDBNull("TestTime") ? null : (DateTime?)reader.GetDateTime("TestTime"),
                                    TestResult = reader.IsDBNull("TestResult") ? null : reader.GetString("TestResult"),
                                    StatusC = reader.IsDBNull("StatusC") ? null : reader.GetString("StatusC"),
                                    E2Record = reader.IsDBNull("E2Record") ? null : reader.GetString("E2Record"),

                                    InVSWR1 = reader.IsDBNull("InVSWR1") ? null : Convert.ToString(reader["InVSWR1"]),
                                    InVSWR2 = reader.IsDBNull("InVSWR2") ? null : Convert.ToString(reader["InVSWR2"]),
                                    InVSWR3 = reader.IsDBNull("InVSWR3") ? null : Convert.ToString(reader["InVSWR3"]),
                                    InVSWR4 = reader.IsDBNull("InVSWR4") ? null : Convert.ToString(reader["InVSWR4"]),

                                    OutVSWR1 = reader.IsDBNull("OutVSWR1") ? null : Convert.ToString(reader["OutVSWR1"]),
                                    OutVSWR2 = reader.IsDBNull("OutVSWR2") ? null : Convert.ToString(reader["OutVSWR2"]),
                                    OutVSWR3 = reader.IsDBNull("OutVSWR3") ? null : Convert.ToString(reader["OutVSWR3"]),
                                    OutVSWR4 = reader.IsDBNull("OutVSWR4") ? null : Convert.ToString(reader["OutVSWR4"]),

                                    Gain1 = reader.IsDBNull("Gain1") ? null : Convert.ToString(reader["Gain1"]),
                                    Gain2 = reader.IsDBNull("Gain2") ? null : Convert.ToString(reader["Gain2"]),
                                    Gain3 = reader.IsDBNull("Gain3") ? null : Convert.ToString(reader["Gain3"]),
                                    Gain4 = reader.IsDBNull("Gain4") ? null : Convert.ToString(reader["Gain4"]),

                                    GainMax1 = reader.IsDBNull("GainMax1") ? null : Convert.ToString(reader["GainMax1"]),
                                    GainMax2 = reader.IsDBNull("GainMax2") ? null : Convert.ToString(reader["GainMax2"]),
                                    GainMax3 = reader.IsDBNull("GainMax3") ? null : Convert.ToString(reader["GainMax3"]),
                                    GainMax4 = reader.IsDBNull("GainMax4") ? null : Convert.ToString(reader["GainMax4"]),

                                    GainIm1 = reader.IsDBNull("GainIm1") ? null : Convert.ToString(reader["GainIm1"]),
                                    GainIm2 = reader.IsDBNull("GainIm2") ? null : Convert.ToString(reader["GainIm2"]),
                                    GainIm3 = reader.IsDBNull("GainIm3") ? null : Convert.ToString(reader["GainIm3"]),
                                    GainIm4 = reader.IsDBNull("GainIm4") ? null : Convert.ToString(reader["GainIm4"]),

                                    Isolation1 = reader.IsDBNull("Isolation1") ? null : Convert.ToString(reader["Isolation1"]),
                                    Isolation2 = reader.IsDBNull("Isolation2") ? null : Convert.ToString(reader["Isolation2"]),
                                    Isolation3 = reader.IsDBNull("Isolation3") ? null : Convert.ToString(reader["Isolation3"]),
                                    Isolation4 = reader.IsDBNull("Isolation4") ? null : Convert.ToString(reader["Isolation4"]),

                                    NF1 = reader.IsDBNull("NF1") ? null : Convert.ToString(reader["NF1"]),
                                    NF2 = reader.IsDBNull("NF2") ? null : Convert.ToString(reader["NF2"]),
                                    NF3 = reader.IsDBNull("NF3") ? null : Convert.ToString(reader["NF3"]),
                                    NF4 = reader.IsDBNull("NF4") ? null : Convert.ToString(reader["NF4"]),

                                    OIP31 = reader.IsDBNull("OIP31") ? null : Convert.ToString(reader["OIP31"]),
                                    OIP32 = reader.IsDBNull("OIP32") ? null : Convert.ToString(reader["OIP32"]),
                                    OIP33 = reader.IsDBNull("OIP33") ? null : Convert.ToString(reader["OIP33"]),
                                    OIP34 = reader.IsDBNull("OIP34") ? null : Convert.ToString(reader["OIP34"]),

                                    LIM31 = reader.IsDBNull("LIM31") ? null : Convert.ToString(reader["LIM31"]),
                                    LIM32 = reader.IsDBNull("LIM32") ? null : Convert.ToString(reader["LIM32"]),
                                    LIM33 = reader.IsDBNull("LIM33") ? null : Convert.ToString(reader["LIM33"]),
                                    LIM34 = reader.IsDBNull("LIM34") ? null : Convert.ToString(reader["LIM34"]),

                                    RIM31 = reader.IsDBNull("RIM31") ? null : Convert.ToString(reader["RIM31"]),
                                    RIM32 = reader.IsDBNull("RIM32") ? null : Convert.ToString(reader["RIM32"]),
                                    RIM33 = reader.IsDBNull("RIM33") ? null : Convert.ToString(reader["RIM33"]),
                                    RIM34 = reader.IsDBNull("RIM34") ? null : Convert.ToString(reader["RIM34"]),

                                    P1dB1 = reader.IsDBNull("P1dB1") ? null : Convert.ToString(reader["P1dB1"]),
                                    P1dB2 = reader.IsDBNull("P1dB2") ? null : Convert.ToString(reader["P1dB2"]),
                                    P1dB3 = reader.IsDBNull("P1dB3") ? null : Convert.ToString(reader["P1dB3"]),
                                    P1dB4 = reader.IsDBNull("P1dB4") ? null : Convert.ToString(reader["P1dB4"]),

                                    Psat1 = reader.IsDBNull("Psat1") ? null : Convert.ToString(reader["Psat1"]),
                                    Psat2 = reader.IsDBNull("Psat2") ? null : Convert.ToString(reader["Psat2"]),
                                    Psat3 = reader.IsDBNull("Psat3") ? null : Convert.ToString(reader["Psat3"]),
                                    Psat4 = reader.IsDBNull("Psat4") ? null : Convert.ToString(reader["Psat4"]),

                                    Current1 = reader.IsDBNull("Current1") ? null : Convert.ToString(reader["Current1"]),
                                    Current2 = reader.IsDBNull("Current2") ? null : Convert.ToString(reader["Current2"]),
                                    Current3 = reader.IsDBNull("Current3") ? null : Convert.ToString(reader["Current3"]),
                                    Current4 = reader.IsDBNull("Current4") ? null : Convert.ToString(reader["Current4"]),

                                    SoftVersion = reader.IsDBNull("SoftVersion") ? null : reader.GetString("SoftVersion")
                                });

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
    




    }

}