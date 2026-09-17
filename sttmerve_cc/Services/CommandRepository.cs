using System;
using System.Collections.Generic;
using System.IO;
using Dapper;
using Microsoft.Data.Sqlite;
using sttmerve_cc.Models;

namespace sttmerve_cc.Services
{
    public class CommandRepository
    {
        private readonly string _connectionString;

        public CommandRepository()
        {
            string dbPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "stt_data.db");
            _connectionString = $"Data Source={dbPath}";
            InitializeDatabase();
        }

        private void InitializeDatabase()
        {
            using var connection = new SqliteConnection(_connectionString);
            connection.Open();
            string sql = @"
                CREATE TABLE IF NOT EXISTS Commands (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    Category TEXT NOT NULL,
                    Type TEXT NOT NULL,
                    Value TEXT NOT NULL,
                    CreatedAt DATETIME NOT NULL
                );";
            connection.Execute(sql);
        }

        public void SaveCommand(CommandPayload payload)
        {
            using var connection = new SqliteConnection(_connectionString);
            string sql = @"
                INSERT INTO Commands (Category, Type, Value, CreatedAt)
                VALUES (@Category, @Type, @Value, @CreatedAt);";

            connection.Execute(sql, new
            {
                payload.Category,
                payload.Type,
                payload.Value,
                CreatedAt = DateTime.Now
            });
        }

        public IEnumerable<CommandPayload> GetAll()
        {
            using var connection = new SqliteConnection(_connectionString);
            return connection.Query<CommandPayload>("SELECT Category, Type, Value FROM Commands ORDER BY CreatedAt ASC;");
        }

        public void ClearAll()
        {
            using var connection = new SqliteConnection(_connectionString);
            connection.Execute("DELETE FROM Commands;");
        }
    }
}
