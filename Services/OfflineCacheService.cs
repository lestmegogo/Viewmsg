using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Data.Sqlite;
using MsgViewer.Models;

namespace MsgViewer.Services
{
    public static class OfflineCacheService
    {
        private static readonly string DbPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "MsgViewer",
            "offline_cache.db");

        static OfflineCacheService()
        {
            try
            {
                string dir = Path.GetDirectoryName(DbPath)!;
                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }

                using (var conn = new SqliteConnection($"Data Source={DbPath}"))
                {
                    conn.Open();
                    
                    // Create Email table
                    string createEmailTable = @"
                        CREATE TABLE IF NOT EXISTS Emails (
                            FilePath TEXT PRIMARY KEY,
                            Subject TEXT,
                            FromName TEXT,
                            FromEmail TEXT,
                            [To] TEXT,
                            Cc TEXT,
                            Date TEXT,
                            IsRead INTEGER,
                            BodyHtml TEXT,
                            BodyText TEXT,
                            FolderId TEXT
                        );";
                    using (var cmd = new SqliteCommand(createEmailTable, conn))
                    {
                        cmd.ExecuteNonQuery();
                    }

                    // Create Attachment table
                    string createAttachmentTable = @"
                        CREATE TABLE IF NOT EXISTS Attachments (
                            Id INTEGER PRIMARY KEY AUTOINCREMENT,
                            MessageFilePath TEXT,
                            FileName TEXT,
                            Data BLOB,
                            FOREIGN KEY(MessageFilePath) REFERENCES Emails(FilePath) ON DELETE CASCADE
                        );";
                    using (var cmd = new SqliteCommand(createAttachmentTable, conn))
                    {
                        cmd.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to initialize SQLite cache db: {ex.Message}");
            }
        }

        public static void SaveEmails(List<EmailMessage> emails, string folderId)
        {
            try
            {
                using (var conn = new SqliteConnection($"Data Source={DbPath}"))
                {
                    conn.Open();
                    using (var transaction = conn.BeginTransaction())
                    {
                        foreach (var email in emails)
                        {
                            string insertEmail = @"
                                INSERT OR REPLACE INTO Emails (FilePath, Subject, FromName, FromEmail, [To], Cc, Date, IsRead, BodyHtml, BodyText, FolderId)
                                VALUES ($path, $subject, $fromName, $fromEmail, $to, $cc, $date, $isRead, $bodyHtml, $bodyText, $folderId);";
                            
                            using (var cmd = new SqliteCommand(insertEmail, conn, transaction))
                            {
                                cmd.Parameters.AddWithValue("$path", email.FilePath);
                                cmd.Parameters.AddWithValue("$subject", email.Subject ?? "");
                                cmd.Parameters.AddWithValue("$fromName", email.FromName ?? "");
                                cmd.Parameters.AddWithValue("$fromEmail", email.FromEmail ?? "");
                                cmd.Parameters.AddWithValue("$to", email.To ?? "");
                                cmd.Parameters.AddWithValue("$cc", email.Cc ?? "");
                                cmd.Parameters.AddWithValue("$date", email.Date?.ToString("o") ?? "");
                                cmd.Parameters.AddWithValue("$isRead", email.IsRead ? 1 : 0);
                                cmd.Parameters.AddWithValue("$bodyHtml", email.BodyHtml ?? "");
                                cmd.Parameters.AddWithValue("$bodyText", email.BodyText ?? "");
                                cmd.Parameters.AddWithValue("$folderId", folderId);
                                cmd.ExecuteNonQuery();
                            }

                            // Save attachments if loaded
                            if (email.Attachments != null && email.Attachments.Count > 0)
                            {
                                // Delete existing attachments first
                                string deleteAtts = "DELETE FROM Attachments WHERE MessageFilePath = $path;";
                                using (var cmd = new SqliteCommand(deleteAtts, conn, transaction))
                                {
                                    cmd.Parameters.AddWithValue("$path", email.FilePath);
                                    cmd.ExecuteNonQuery();
                                }

                                foreach (var att in email.Attachments)
                                {
                                    string insertAtt = "INSERT INTO Attachments (MessageFilePath, FileName, Data) VALUES ($path, $name, $data);";
                                    using (var cmd = new SqliteCommand(insertAtt, conn, transaction))
                                    {
                                        cmd.Parameters.AddWithValue("$path", email.FilePath);
                                        cmd.Parameters.AddWithValue("$name", att.FileName ?? "attachment");
                                        cmd.Parameters.AddWithValue("$data", att.Data ?? Array.Empty<byte>());
                                        cmd.ExecuteNonQuery();
                                    }
                                }
                            }
                        }
                        transaction.Commit();
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to save emails to SQLite cache: {ex.Message}");
            }
        }

        public static List<EmailMessage> LoadEmails(string folderId)
        {
            var emails = new List<EmailMessage>();
            try
            {
                using (var conn = new SqliteConnection($"Data Source={DbPath}"))
                {
                    conn.Open();
                    string query = "SELECT FilePath, Subject, FromName, FromEmail, [To], Cc, Date, IsRead, BodyHtml, BodyText FROM Emails WHERE FolderId = $folderId;";
                    using (var cmd = new SqliteCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("$folderId", folderId);
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string filePath = reader.GetString(0);
                                string subject = reader.GetString(1);
                                string fromName = reader.GetString(2);
                                string fromEmail = reader.GetString(3);
                                string to = reader.GetString(4);
                                string cc = reader.GetString(5);
                                string dateStr = reader.GetString(6);
                                bool isRead = reader.GetInt32(7) == 1;
                                string bodyHtml = reader.IsDBNull(8) ? "" : reader.GetString(8);
                                string bodyText = reader.IsDBNull(9) ? "" : reader.GetString(9);

                                DateTime? date = null;
                                if (DateTime.TryParse(dateStr, out var d)) date = d;

                                var email = new EmailMessage
                                {
                                    FilePath = filePath,
                                    Subject = subject,
                                    FromName = fromName,
                                    FromEmail = fromEmail,
                                    To = to,
                                    Cc = cc,
                                    Date = date,
                                    IsRead = isRead,
                                    BodyHtml = bodyHtml,
                                    BodyText = bodyText
                                };

                                // Load attachments from DB
                                LoadAttachmentsForMessage(email, conn);

                                emails.Add(email);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to load emails from SQLite cache: {ex.Message}");
            }
            return emails;
        }

        private static void LoadAttachmentsForMessage(EmailMessage email, SqliteConnection conn)
        {
            try
            {
                string query = "SELECT FileName, Data FROM Attachments WHERE MessageFilePath = $path;";
                using (var cmd = new SqliteCommand(query, conn))
                {
                    cmd.Parameters.AddWithValue("$path", email.FilePath);
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string fileName = reader.GetString(0);
                            byte[] data = (byte[])reader.GetValue(1);
                            email.Attachments.Add(new EmailAttachment
                            {
                                FileName = fileName,
                                Data = data
                            });
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to load attachments for {email.Subject}: {ex.Message}");
            }
        }

        public static void ClearCache()
        {
            try
            {
                if (File.Exists(DbPath))
                {
                    File.Delete(DbPath);
                }
            }
            catch {}
        }
    }
}
