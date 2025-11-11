using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using Dapper;
using System.IO;



namespace ChipManualGenerationSogt
{
    internal class SqlServerDBAccess
    {
    }

    // 用Dapper 一定要注意， 数据结构和数据库字段名字 类型都要对应，否则会出现类型转换异常
    public class UserRepository
    {
        private readonly string _connectionString;

        public UserRepository(string connectionString)
        {
            _connectionString = connectionString;
        }

        public async Task<List<User>> GetAllUsersAsync()
        {
            var connection = new SqlConnection(_connectionString);
            return (await connection.QueryAsync<User>(
                "SELECT * FROM yp_test_user"
            )).ToList();
        }
    }

    public class TaskRepository
    {

        // private readonly string _connectionString =  "Server=192.168.1.77,1433;Database=QotanaTestSystem;User ID=sa;Password=123456;Encrypt=false;TrustServerCertificate=true;";
        private readonly string _connectionString = "Server=192.168.1.209;Database=mlChips;User ID=sa;Password=qotana;Encrypt=false;TrustServerCertificate=true;";


        public TaskRepository(string connectionString)
        {
            _connectionString = connectionString;
        }
        public TaskRepository()
        {
           
        }
        public async Task<List<TaskModel_DB>> GetAllFinishedTasksAsync()
        {
            var connection = new SqlConnection(_connectionString);
            return (await connection.QueryAsync<TaskModel_DB>(
                "SELECT * FROM yp_test_finished_task"
            )).ToList();
        }

        public async Task<TaskModel_DB> GetSigleFinishedTasksAsync_ID(int id)
        {
            var connection = new SqlConnection(_connectionString);
            return (await connection.QueryFirstOrDefaultAsync<TaskModel_DB>(
                "SELECT * FROM yp_test_finished_task    WHERE ID=@ID", new { ID = id }
            ));
        }

        public async Task<bool> UpdateFinishedTaskAsync(TaskModel_DB task)
        {
            const string sql = @"
                    UPDATE yp_test_finished_task 
             SET 
            TaskName = @TaskName,
            Status = @Status,
            StartDate = @StartDate,
            Consumed = @Consumed,
            DataStatus = @DataStatus,
            FilesStatus = @FilesStatus,
            CheckStatus = @CheckStatus
            WHERE ID = @ID";

            using (var connection = new SqlConnection(_connectionString))
            {
                var rowsAffected = await connection.ExecuteAsync(sql, task);
                return rowsAffected > 0;
            }
        }

        public async Task<bool> InsertFinishedTaskAsync(TaskModel_DB task)
        {
          
            const string sql = @"
                    INSERT INTO yp_test_finished_task 
                    (ID, TaskName, Status, StartDate, Consumed, DataStatus, FilesStatus, CheckStatus)
                    VALUES 
                    (@ID, @TaskName, @Status, @StartDate, @Consumed, @DataStatus, @FilesStatus, @CheckStatus);
                    SELECT @ID;";

            using (var connection = new SqlConnection(_connectionString))
            {
                var newId = await connection.QuerySingleAsync<int>(sql, task);
                return newId >0;
            }


        }


        public async Task<List<TaskModel_DB>> GetAllCurrentTasksAsync()
        {
            var connection = new SqlConnection(_connectionString);
            return (await connection.QueryAsync<TaskModel_DB>(
                "SELECT * FROM yp_test_current_task"
            )).ToList();
        }

        public async Task<TaskModel_DB> GetSigleCurrentTasksAsync_ID(int id)
        {
            var connection = new SqlConnection(_connectionString);
            return (await connection.QueryFirstOrDefaultAsync<TaskModel_DB>(
                "SELECT * FROM yp_test_current_task    WHERE ID=@ID", new { ID = id }
            ));
        }

        public async Task<TaskModel_DB> GetLastCurrentTasksAsync()
        {
            const string sql = @"
                    SELECT TOP 1 * 
                    FROM yp_test_current_task 
                    ORDER BY ID DESC";
            using (var connection = new SqlConnection(_connectionString))
            {
                return await connection.QueryFirstOrDefaultAsync<TaskModel_DB>(sql);
            }
        }


        public async Task<bool> UpdateCurrentTaskAsync(TaskModel_DB task)
        {
            const string sql = @"
                    UPDATE yp_test_current_task 
            SET 
            TaskName = @TaskName,
            Status = @Status,
            StartDate = @StartDate,
            Consumed = @Consumed,
            DataStatus = @DataStatus,
            FilesStatus = @FilesStatus,
            CheckStatus = @CheckStatus
            WHERE ID = @ID";

            using (var connection = new SqlConnection(_connectionString))
            {
                var rowsAffected = await connection.ExecuteAsync(sql, task);
                return rowsAffected > 0;
            }
        }

        public async Task<int> InsertCurrentTaskAsync(TaskModel_DB task)
        {
            const string sql = @"
                    INSERT INTO yp_test_current_task 
                    (ID, TaskName, Status, StartDate, Consumed, DataStatus, FilesStatus, CheckStatus)
                    VALUES 
                    (@ID, @TaskName, @Status, @StartDate, @Consumed, @DataStatus, @FilesStatus, @CheckStatus);
                    SELECT @ID;";


            using (var connection = new SqlConnection(_connectionString))
            {
                if (task.ID <= 0)
                    throw new ArgumentException("ID Must be greater than 0", nameof(task.ID));

                var newId = await connection.QuerySingleAsync<int>(sql, task);
                return newId;
            }


        }


        /// <summary>
        /// 根据任务ID删除 'yp_test_current_task' 表中的一条记录。
        /// </summary>
        /// <param name="task">包含要删除记录ID的任务模型。</param>
        /// <returns>如果至少删除了一行，则返回 true。</returns>
        public async Task<bool> DeleteCurrentTaskAsync(TaskModel_DB task)
        {
            // 核心变动：使用 DELETE 语句，并只保留 WHERE ID = @ID
            const string sql = @"
            DELETE FROM yp_test_current_task
            WHERE ID = @ID"; // 只根据 ID 来匹配并删除记录

            using (var connection = new SqlConnection(_connectionString))
            {
                // ExecuteAsync 方法会执行 SQL 语句并返回受影响的行数
                var rowsAffected = await connection.ExecuteAsync(sql, task);

                // 如果 rowsAffected 大于 0，表示至少有一条记录被删除
                return rowsAffected > 0;
            }
        }



        /// <summary>
        /// 将单个日志记录异步插入到数据库中。
        /// </summary>
        /// <param name="log">要插入的日志模型对象。</param>
        /// <returns>如果插入成功，则返回 true。</returns>
        public async Task<bool> InsertLogAsync(LogModel log)
        {
            // 自动映射 LogModel 的属性到 SQL 参数
            const string sql = @"
            INSERT INTO  logs(
                TimeStamp, UserName, TaskId, TaskName, 
                Level, Message, PN, SN, ChipNumber
            )
            VALUES (
                @TimeStamp, @UserName, @TaskId, @TaskName, 
                @Level, @Message, @PN, @SN, @ChipNumber
            )";

            using (var connection = new SqlConnection(_connectionString))
            {
                // ExecuteAsync 执行非查询操作，返回受影响的行数
                var rowsAffected = await connection.ExecuteAsync(sql, log);
                return rowsAffected > 0;
            }
        }


        /// <summary>
        /// 异步获取数据库中的所有日志记录。
        /// </summary>
        /// <returns>日志模型列表。</returns>
        public async Task<List<LogModel>> GetAllLogsAsync()
        {
            const string sql = @"
            SELECT *

            FROM  logs
            ORDER BY TimeStamp DESC"; // 按时间戳降序排序

            using (var connection = new SqlConnection(_connectionString))
            {
                // QueryAsync 执行查询操作，并自动映射到 LogModel 列表
                var logs = await connection.QueryAsync<LogModel>(sql);
                return logs.AsList();
            }
        }

        public async Task<List<LogModel>> QueryLogsByUserAsync(
        string userName)
        {
            const string sql = @"
            SELECT *
   
            FROM logs
            WHERE 
                UserName = @UserName
            ORDER BY TimeStamp DESC";

            using (var connection = new SqlConnection(_connectionString))
            {
                var logs = await connection.QueryAsync<LogModel>(sql, new
                {
                    UserName = userName,
                });
                return logs.AsList();
            }
        }

        public async Task<List<LogModel>> QueryLogsAsync(
    string userName = null,
    string taskId = null,
    string level = null,
    string pn = null,
    string sn = null,
    string chipNumber = null,
    DateTime? startTime = null,
    DateTime? endTime = null)
        {
            var sql = new StringBuilder(@"
        SELECT *
        FROM logs
        WHERE 1=1
    ");

            var parameters = new DynamicParameters();

            if (!string.IsNullOrWhiteSpace(userName))
            {
                sql.Append(" AND UserName = @UserName");
                parameters.Add("UserName", userName);
            }

            if (!string.IsNullOrWhiteSpace(taskId))
            {
                sql.Append(" AND TaskId = @TaskId");
                parameters.Add("TaskId", taskId);
            }

            if (!string.IsNullOrWhiteSpace(level))
            {
                sql.Append(" AND Level = @Level");
                parameters.Add("Level", level);
            }

            // 模糊查询部分
            if (!string.IsNullOrWhiteSpace(pn))
            {
                sql.Append(" AND PN LIKE @PN");
                parameters.Add("PN", $"%{pn}%");
            }

            if (!string.IsNullOrWhiteSpace(sn))
            {
                sql.Append(" AND SN LIKE @SN");
                parameters.Add("SN", $"%{sn}%");
            }

            if (!string.IsNullOrWhiteSpace(chipNumber))
            {
                sql.Append(" AND ChipNumber LIKE @ChipNumber");
                parameters.Add("ChipNumber", $"%{chipNumber}%");
            }

            if (startTime.HasValue)
            {
                sql.Append(" AND TimeStamp >= @StartTime");
                parameters.Add("StartTime", startTime.Value);
            }

            if (endTime.HasValue)
            {
                sql.Append(" AND TimeStamp <= @EndTime");
                parameters.Add("EndTime", endTime.Value);
            }

            sql.Append(" ORDER BY TimeStamp DESC");

            using (var connection = new SqlConnection(_connectionString))
            {
                var logs = await connection.QueryAsync<LogModel>(sql.ToString(), parameters);
                return logs.AsList();
            }
        }



        /// <summary>
        /// 异步添加一条新的 OperationModel 记录到数据库。
        /// </summary>
        /// <param name="operation">要插入的 OperationModel 对象。</param>
        /// <returns>插入成功返回 true，否则返回 false。</returns>
        public async Task<bool> InsertOperationAsync(OperationModel operation)
        {
            // 假设 TaskID 是由 OperationModel 传入的，需要手动设置，而不是自增。
            const string sql = @"
        INSERT INTO yp_test_taskcore (
            TaskID,
            TaskName, TimeStamp, StartDateTime, EndDateTime, PN, SN, 
            DataReady, Condition, Data, SourceFiles, FileReady, PptPath
        )
        VALUES (
            @TaskID, 
            @TaskName, @TimeStamp, @StartDateTime, @EndDateTime, @PN, @SN, 
            @DataReady, @Condition, @Data, @SourceFiles, @FileReady, @PptPath
        )";

            using (var connection = new SqlConnection(_connectionString))
            {
                // ExecuteAsync 返回受影响的行数，如果成功，rowsAffected > 0
                // Dapper 会自动将 operation 对象的 TaskID 属性映射到 @TaskID 参数
                int rowsAffected = await connection.ExecuteAsync(sql, operation);

                // 不需要返回 ID，只需返回操作是否成功
                return rowsAffected > 0;
            }
        }


        /// <summary>
        /// 异步按 TaskID 查询一条 OperationModel 记录。
        /// </summary>
        /// <param name="taskId">要查询的记录的 TaskID。</param>
        /// <returns>找到返回 OperationModel，否则返回 null。</returns>
        public async Task<OperationModel> GetOperationByIdAsync(int taskId)
        {
            const string sql = @"
            SELECT *
            FROM yp_test_taskcore
            WHERE TaskID = @TaskID";

            using (var connection = new SqlConnection(_connectionString))
            {
                // QueryFirstOrDefaultAsync 适用于查询单个对象
                return await connection.QueryFirstOrDefaultAsync<OperationModel>(
                    sql,
                    new { TaskID = taskId }
                );
            }
        }



        /// <summary>
        /// 异步更新一条现有的 OperationModel 记录。
        /// </summary>
        /// <param name="operation">包含 TaskID 和新值的 OperationModel 对象。</param>
        /// <returns>更新成功返回 true，否则返回 false。</returns>
        public async Task<bool> UpdateOperationAsync(OperationModel operation)
        {
            // 排除 TaskID 以外的所有字段都应该被更新
            const string sql = @"
             UPDATE yp_test_taskcore SET
                 TaskName = @TaskName,
                 TimeStamp = @TimeStamp,
                 StartDateTime = @StartDateTime,
                 EndDateTime = @EndDateTime,
                 PN = @PN,
                 SN = @SN,
                 DataReady = @DataReady,
                 Condition = @Condition,
                 Data = @Data,
                 SourceFiles = @SourceFiles,
                 FileReady = @FileReady,
                 PptPath = @PptPath
             WHERE TaskID = @TaskID"; // WHERE 子句使用 TaskID 进行匹配

            using (var connection = new SqlConnection(_connectionString))
            {
                int rowsAffected = await connection.ExecuteAsync(sql, operation);
                return rowsAffected > 0;
            }
        }   
    }


    public class TaskSqlServerRepository
    {
        private readonly string _connectionString = "Server=192.168.1.209;Database=mlChips;User ID=sa;Password=qotana;Encrypt=false;TrustServerCertificate=true;";
        private const string TableName = "tasks"; // 替换为你的实际表名

        private IDbConnection CreateConnection()
        {
            return new SqlConnection(_connectionString);
        }
        // ==========================================================
        // C - Create (异步增加)
        // ==========================================================
        public async Task<int> AddTaskAsync(TaskSqlServerModel task)
        {
            var sql = $@"
            INSERT INTO {TableName} 
            (ID,PPTModel, TaskName,PptName, Status, Level, Major, Minor, StartDate, EndDate, DataStatus, FilesStatus, Conditions, TableUpdate)
            VALUES 
            (@ID,@PPTModel, @TaskName,@PptName, @Status, @Level, @Major, @Minor, @StartDate, @EndDate, @DataStatus, @FilesStatus, @Conditions,@TableUpdate);
            SELECT CAST(SCOPE_IDENTITY() as int)";

            using (var connection = CreateConnection())
            {
                // 使用 ExecuteScalarAsync<int> 异步执行插入并返回新 ID
                return await connection.ExecuteScalarAsync<int>(sql, task);
            }
        }

        // ==========================================================
        // R - Read (异步查询)
        // ==========================================================

        // 1. 异步查询所有记录
        public async Task<IEnumerable<TaskSqlServerModel>> GetAllTasksAsync()
        {
            var sql = $"SELECT * FROM {TableName}";
            using (var connection = CreateConnection())
            {
                // 使用 QueryAsync<T> 异步返回所有记录的列表
                return await connection.QueryAsync<TaskSqlServerModel>(sql);
            }
        }

        // 2. 根据 ID 异步查询单条记录
        public async Task<TaskSqlServerModel> GetTaskByIdAsync(int id)
        {
            var sql = $"SELECT * FROM {TableName} WHERE ID = @Id";
            using (var connection = CreateConnection())
            {
                // 使用 QueryFirstOrDefaultAsync<T> 异步返回第一个匹配项
                return await connection.QueryFirstOrDefaultAsync<TaskSqlServerModel>(sql, new { ID = id });
            }
        }

        // ==========================================================
        // U - Update (异步更新)
        // ==========================================================
        public async Task<bool> UpdateTaskAsync(TaskSqlServerModel task)
        {
            var sql = $@"
            UPDATE {TableName} SET
                PPTModel = @PPTModel,
                TaskName = @TaskName,
                Status = @Status,
                Level = @Level,
                Major = @Major,
                Minor = @Minor,
                StartDate = @StartDate,
                EndDate = @EndDate,
                DataStatus = @DataStatus,
                FilesStatus = @FilesStatus,
                Conditions = @Conditions,
                PptName = @PptName,
                TableUpdate = @TableUpdate
            WHERE ID = @ID";

            using (var connection = CreateConnection())
            {
                // 使用 ExecuteAsync 异步执行更新，并返回受影响的行数
                return await connection.ExecuteAsync(sql, task) > 0;
            }
        }

        // ==========================================================
        // D - Delete (异步删除)
        // ==========================================================
        public async Task<bool> DeleteTaskAsync(int id)
        {
            var sql = $"DELETE FROM {TableName} WHERE ID = @Id";

            using (var connection = CreateConnection())
            {
                // 使用 ExecuteAsync 异步执行删除
                return await connection.ExecuteAsync(sql, new { ID = id }) > 0;
            }
        }
    }


    public class LogRepository
    {
        public static string _connectionString = "Server=192.168.1.209;Database=mlChips;User ID=sa;Password=qotana;Encrypt=false;TrustServerCertificate=true;";
        private const string TableName = "logs"; // 替换为你的实际表名

        static public async Task<bool> InsertLogAsync(LogModel log)
        {
            // 自动映射 LogModel 的属性到 SQL 参数
            const string sql = @"
            INSERT INTO  logs(
                TimeStamp, UserName, TaskId, TaskName, 
                Level, Message, PN, SN, ChipNumber
            )
            VALUES (
                @TimeStamp, @UserName, @TaskId, @TaskName, 
                @Level, @Message, @PN, @SN, @ChipNumber
            )";

            using (var connection = new SqlConnection(_connectionString))
            {
                // ExecuteAsync 执行非查询操作，返回受影响的行数
                var rowsAffected = await connection.ExecuteAsync(sql, log);
                return rowsAffected > 0;
            }
        }


        /// <summary>
        /// 异步获取数据库中的所有日志记录。
        /// </summary>
        /// <returns>日志模型列表。</returns>
        static public async Task<List<LogModel>> GetAllLogsAsync()
        {
            const string sql = @"
            SELECT *

            FROM  logs
            ORDER BY TimeStamp DESC"; // 按时间戳降序排序

            using (var connection = new SqlConnection(_connectionString))
            {
                // QueryAsync 执行查询操作，并自动映射到 LogModel 列表
                var logs = await connection.QueryAsync<LogModel>(sql);
                return logs.AsList();
            }
        }

  
       static public async Task<List<LogModel>> QueryLogsAsync(
    string userName = null,
    string taskId = null,
    string level = null,
    string pn = null,
    string sn = null,
    string chipNumber = null,
    DateTime? startTime = null,
    DateTime? endTime = null)
        {
            var sql = new StringBuilder(@"
        SELECT *
        FROM logs
        WHERE 1=1
    ");

            var parameters = new DynamicParameters();

            if (!string.IsNullOrWhiteSpace(userName))
            {
                sql.Append(" AND UserName = @UserName");
                parameters.Add("UserName", userName);
            }

            if (!string.IsNullOrWhiteSpace(taskId))
            {
                sql.Append(" AND TaskId = @TaskId");
                parameters.Add("TaskId", taskId);
            }

            if (!string.IsNullOrWhiteSpace(level))
            {
                sql.Append(" AND Level = @Level");
                parameters.Add("Level", level);
            }

            // 模糊查询部分
            if (!string.IsNullOrWhiteSpace(pn))
            {
                sql.Append(" AND PN LIKE @PN");
                parameters.Add("PN", $"%{pn}%");
            }

            if (!string.IsNullOrWhiteSpace(sn))
            {
                sql.Append(" AND SN LIKE @SN");
                parameters.Add("SN", $"%{sn}%");
            }

            if (!string.IsNullOrWhiteSpace(chipNumber))
            {
                sql.Append(" AND ChipNumber LIKE @ChipNumber");
                parameters.Add("ChipNumber", $"%{chipNumber}%");
            }

            if (startTime.HasValue)
            {
                sql.Append(" AND TimeStamp >= @StartTime");
                parameters.Add("StartTime", startTime.Value);
            }

            if (endTime.HasValue)
            {
                sql.Append(" AND TimeStamp <= @EndTime");
                parameters.Add("EndTime", endTime.Value);
            }

            sql.Append(" ORDER BY TimeStamp DESC");

            using (var connection = new SqlConnection(_connectionString))
            {
                var logs = await connection.QueryAsync<LogModel>(sql.ToString(), parameters);
                return logs.AsList();
            }
        }




    }

    /// <summary>
    /// 对应 JSON 结构中的 "DatabaseConnection" 对象。
    /// </summary>





   //通用数据访问层 + 业务逻辑分离 适合表很多的时候
    //public class DapperRepository
    //{
    //    private readonly string _connectionString;

    //    public DapperRepository(string connectionString)
    //    {
    //        _connectionString = connectionString
    //            ?? throw new ArgumentNullException(nameof(connectionString));
    //    }

    //    // ======================
    //    // 多行查询 - 同步
    //    // ======================
    //    public IEnumerable<T> Query<T>(string sql, object? param = null)
    //    {
    //        using var connection = new SqlConnection(_connectionString);
    //        return connection.Query<T>(sql, param);
    //    }

    //    // ======================
    //    // 多行查询 - 异步（推荐用于 Web/API）
    //    // ======================
    //    public async Task<IEnumerable<T>> QueryAsync<T>(string sql, object? param = null)
    //    {
    //        using var connection = new SqlConnection(_connectionString);
    //        return await connection.QueryAsync<T>(sql, param);
    //    }

    //    // ======================
    //    // 单行查询（可能为 null）- 同步
    //    // ======================
    //    public T? QueryFirstOrDefault<T>(string sql, object? param = null) where T : class
    //    {
    //        using var connection = new SqlConnection(_connectionString);
    //        return connection.QueryFirstOrDefault<T>(sql, param);
    //    }

    //    public T QueryFirstOrDefault<T>(string sql, object? param = null) where T : struct
    //    {
    //        using var connection = new SqlConnection(_connectionString);
    //        return connection.QueryFirstOrDefault<T>(sql, param);
    //    }

    //    // ======================
    //    // 单行查询 - 异步
    //    // ======================
    //    public async Task<T?> QueryFirstOrDefaultAsync<T>(string sql, object? param = null) where T : class
    //    {
    //        using var connection = new SqlConnection(_connectionString);
    //        return await connection.QueryFirstOrDefaultAsync<T>(sql, param);
    //    }

    //    public async Task<T> QueryFirstOrDefaultAsync<T>(string sql, object? param = null) where T : struct
    //    {
    //        using var connection = new SqlConnection(_connectionString);
    //        return await connection.QueryFirstOrDefaultAsync<T>(sql, param);
    //    }

    //    // ======================
    //    // 执行非查询（INSERT/UPDATE/DELETE）- 同步
    //    // ======================
    //    public int Execute(string sql, object? param = null)
    //    {
    //        using var connection = new SqlConnection(_connectionString);
    //        return connection.Execute(sql, param);
    //    }

    //    // ======================
    //    // 执行非查询 - 异步
    //    // ======================
    //    public async Task<int> ExecuteAsync(string sql, object? param = null)
    //    {
    //        using var connection = new SqlConnection(_connectionString);
    //        return await connection.ExecuteAsync(sql, param);
    //    }
    //}

}


