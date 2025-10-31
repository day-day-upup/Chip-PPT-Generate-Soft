using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Dapper;



namespace ChipManualGenerationSogt
{
    internal class SqlServerDBAccess
    {
    }

    // ��Dapper һ��Ҫע�⣬ ���ݽṹ�����ݿ��ֶ����� ���Ͷ�Ҫ��Ӧ��������������ת���쳣
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

        private readonly string _connectionString =  "Server=192.168.1.77,1433;Database=QotanaTestSystem;User ID=sa;Password=123456;Encrypt=false;TrustServerCertificate=true;";


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
        /// ��������IDɾ�� 'yp_test_current_task' ���е�һ����¼��
        /// </summary>
        /// <param name="task">����Ҫɾ����¼ID������ģ�͡�</param>
        /// <returns>�������ɾ����һ�У��򷵻� true��</returns>
        public async Task<bool> DeleteCurrentTaskAsync(TaskModel_DB task)
        {
            // ���ı䶯��ʹ�� DELETE ��䣬��ֻ���� WHERE ID = @ID
            const string sql = @"
            DELETE FROM yp_test_current_task
            WHERE ID = @ID"; // ֻ���� ID ��ƥ�䲢ɾ����¼

            using (var connection = new SqlConnection(_connectionString))
            {
                // ExecuteAsync ������ִ�� SQL ��䲢������Ӱ�������
                var rowsAffected = await connection.ExecuteAsync(sql, task);

                // ��� rowsAffected ���� 0����ʾ������һ����¼��ɾ��
                return rowsAffected > 0;
            }
        }



        /// <summary>
        /// ��������־��¼�첽���뵽���ݿ��С�
        /// </summary>
        /// <param name="log">Ҫ�������־ģ�Ͷ���</param>
        /// <returns>�������ɹ����򷵻� true��</returns>
        public async Task<bool> InsertLogAsync(LogModel log)
        {
            // �Զ�ӳ�� LogModel �����Ե� SQL ����
            const string sql = @"
            INSERT INTO  yp_test_logs(
                TimeStamp, UserName, TaskId, TaskName, 
                Level, Message, PN, SN, ChipNumber
            )
            VALUES (
                @TimeStamp, @UserName, @TaskId, @TaskName, 
                @Level, @Message, @PN, @SN, @ChipNumber
            )";

            using (var connection = new SqlConnection(_connectionString))
            {
                // ExecuteAsync ִ�зǲ�ѯ������������Ӱ�������
                var rowsAffected = await connection.ExecuteAsync(sql, log);
                return rowsAffected > 0;
            }
        }


        /// <summary>
        /// �첽��ȡ���ݿ��е�������־��¼��
        /// </summary>
        /// <returns>��־ģ���б�</returns>
        public async Task<List<LogModel>> GetAllLogsAsync()
        {
            const string sql = @"
            SELECT *

            FROM  yp_test_logs
            ORDER BY TimeStamp DESC"; // ��ʱ�����������

            using (var connection = new SqlConnection(_connectionString))
            {
                // QueryAsync ִ�в�ѯ���������Զ�ӳ�䵽 LogModel �б�
                var logs = await connection.QueryAsync<LogModel>(sql);
                return logs.AsList();
            }
        }

        public async Task<List<LogModel>> QueryLogsByUserAsync(
        string userName)
        {
            const string sql = @"
            SELECT *
   
            FROM yp_test_logs
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
        FROM yp_test_logs
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

            // ģ����ѯ����
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
        /// �첽���һ���µ� OperationModel ��¼�����ݿ⡣
        /// </summary>
        /// <param name="operation">Ҫ����� OperationModel ����</param>
        /// <returns>����ɹ����� true�����򷵻� false��</returns>
        public async Task<bool> InsertOperationAsync(OperationModel operation)
        {
            // ���� TaskID ���� OperationModel ����ģ���Ҫ�ֶ����ã�������������
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
                // ExecuteAsync ������Ӱ�������������ɹ���rowsAffected > 0
                // Dapper ���Զ��� operation ����� TaskID ����ӳ�䵽 @TaskID ����
                int rowsAffected = await connection.ExecuteAsync(sql, operation);

                // ����Ҫ���� ID��ֻ�践�ز����Ƿ�ɹ�
                return rowsAffected > 0;
            }
        }


        /// <summary>
        /// �첽�� TaskID ��ѯһ�� OperationModel ��¼��
        /// </summary>
        /// <param name="taskId">Ҫ��ѯ�ļ�¼�� TaskID��</param>
        /// <returns>�ҵ����� OperationModel�����򷵻� null��</returns>
        public async Task<OperationModel> GetOperationByIdAsync(int taskId)
        {
            const string sql = @"
            SELECT *
            FROM yp_test_taskcore
            WHERE TaskID = @TaskID";

            using (var connection = new SqlConnection(_connectionString))
            {
                // QueryFirstOrDefaultAsync �����ڲ�ѯ��������
                return await connection.QueryFirstOrDefaultAsync<OperationModel>(
                    sql,
                    new { TaskID = taskId }
                );
            }
        }



        /// <summary>
        /// �첽����һ�����е� OperationModel ��¼��
        /// </summary>
        /// <param name="operation">���� TaskID ����ֵ�� OperationModel ����</param>
        /// <returns>���³ɹ����� true�����򷵻� false��</returns>
        public async Task<bool> UpdateOperationAsync(OperationModel operation)
        {
            // �ų� TaskID ����������ֶζ�Ӧ�ñ�����
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
             WHERE TaskID = @TaskID"; // WHERE �Ӿ�ʹ�� TaskID ����ƥ��

            using (var connection = new SqlConnection(_connectionString))
            {
                int rowsAffected = await connection.ExecuteAsync(sql, operation);
                return rowsAffected > 0;
            }
        }   
    }



    //ͨ�����ݷ��ʲ� + ҵ���߼����� �ʺϱ�ܶ��ʱ��
    //public class DapperRepository
    //{
    //    private readonly string _connectionString;

    //    public DapperRepository(string connectionString)
    //    {
    //        _connectionString = connectionString
    //            ?? throw new ArgumentNullException(nameof(connectionString));
    //    }

    //    // ======================
    //    // ���в�ѯ - ͬ��
    //    // ======================
    //    public IEnumerable<T> Query<T>(string sql, object? param = null)
    //    {
    //        using var connection = new SqlConnection(_connectionString);
    //        return connection.Query<T>(sql, param);
    //    }

    //    // ======================
    //    // ���в�ѯ - �첽���Ƽ����� Web/API��
    //    // ======================
    //    public async Task<IEnumerable<T>> QueryAsync<T>(string sql, object? param = null)
    //    {
    //        using var connection = new SqlConnection(_connectionString);
    //        return await connection.QueryAsync<T>(sql, param);
    //    }

    //    // ======================
    //    // ���в�ѯ������Ϊ null��- ͬ��
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
    //    // ���в�ѯ - �첽
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
    //    // ִ�зǲ�ѯ��INSERT/UPDATE/DELETE��- ͬ��
    //    // ======================
    //    public int Execute(string sql, object? param = null)
    //    {
    //        using var connection = new SqlConnection(_connectionString);
    //        return connection.Execute(sql, param);
    //    }

    //    // ======================
    //    // ִ�зǲ�ѯ - �첽
    //    // ======================
    //    public async Task<int> ExecuteAsync(string sql, object? param = null)
    //    {
    //        using var connection = new SqlConnection(_connectionString);
    //        return await connection.ExecuteAsync(sql, param);
    //    }
    //}

}


