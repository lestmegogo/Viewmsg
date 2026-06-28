using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Microsoft.Identity.Client;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Azure.Identity;
using MsgViewer.Models;

namespace MsgViewer.Services
{
    public class Office365Service
    {
        private static string ClientId = "327c9f81-4e1f-46c7-b7d4-9f24a84173b6"; // User registered Client ID
        private static readonly string[] Scopes = { "User.Read", "Mail.ReadWrite", "Mail.Send" };

        private static IPublicClientApplication? _pca;
        private static GraphServiceClient? _graphClient;
        private static IAccount? _currentUserAccount;

        public static string? UserDisplayName { get; private set; }
        public static string? UserEmail { get; private set; }
        public static byte[]? UserAvatar { get; private set; }

        public static bool IsSignedIn => _currentUserAccount != null;

        public static async Task InitializeAsync()
        {
            string configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "o365_config.txt");
            if (File.Exists(configPath))
            {
                try
                {
                    string fileContent = File.ReadAllText(configPath);
                    var lines = fileContent.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                    if (lines.Length > 0)
                    {
                        var guidStr = lines[0].Trim();
                        if (Guid.TryParse(guidStr, out _))
                        {
                            ClientId = guidStr;
                        }
                    }
                }
                catch {}
            }
            else
            {
                try
                {
                    File.WriteAllText(configPath, ClientId + "\r\n# Thay thế dòng 1 bằng Client ID (Application ID) của bạn nếu muốn sử dụng App Registration riêng trên Azure Active Directory.");
                }
                catch {}
            }

            _pca = PublicClientApplicationBuilder.Create(ClientId)
                .WithAuthority(AadAuthorityAudience.AzureAdMultipleOrgs)
                .WithRedirectUri("http://localhost")
                .Build();

            // Try to sign in silently
            var accounts = await _pca.GetAccountsAsync();
            if (accounts.Any())
            {
                try
                {
                    var result = await _pca.AcquireTokenSilent(Scopes, accounts.First()).ExecuteAsync();
                    _currentUserAccount = result.Account;
                    InitializeGraphClient(result.AccessToken);
                    await LoadUserProfileAsync();
                }
                catch (MsalUiRequiredException)
                {
                    // Silent login failed, need explicit login
                    _currentUserAccount = null;
                }
            }
        }

        public static async Task<bool> SignInAsync()
        {
            if (_pca == null) await InitializeAsync();

            try
            {
                var result = await _pca!.AcquireTokenInteractive(Scopes).ExecuteAsync();
                _currentUserAccount = result.Account;
                InitializeGraphClient(result.AccessToken);
                await LoadUserProfileAsync();
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Sign-in failed: {ex.Message}");
                return false;
            }
        }

        public static async Task SignOutAsync()
        {
            if (_pca != null && _currentUserAccount != null)
            {
                await _pca.RemoveAsync(_currentUserAccount);
                _currentUserAccount = null;
                _graphClient = null;
                UserDisplayName = null;
                UserEmail = null;
                UserAvatar = null;
            }
        }

        private static void InitializeGraphClient(string accessToken)
        {
            var tokenCredential = new TokenProviderCredential(() => Task.FromResult(accessToken));
            _graphClient = new GraphServiceClient(tokenCredential);
        }

        private static async Task LoadUserProfileAsync()
        {
            if (_graphClient == null) return;

            try
            {
                var me = await _graphClient.Me.GetAsync();
                UserDisplayName = me?.DisplayName;
                UserEmail = me?.Mail ?? me?.UserPrincipalName;

                // Load avatar photo
                try
                {
                    using (var photoStream = await _graphClient.Me.Photo.Content.GetAsync())
                    {
                        if (photoStream != null)
                        {
                            using (var ms = new MemoryStream())
                            {
                                await photoStream.CopyToAsync(ms);
                                UserAvatar = ms.ToArray();
                            }
                        }
                    }
                }
                catch
                {
                    UserAvatar = null; // No photo or no permission
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to load user profile: {ex.Message}");
            }
        }

        public static async Task<List<Tuple<string, string>>> GetFoldersAsync()
        {
            var foldersList = new List<Tuple<string, string>>();
            if (_graphClient == null) return foldersList;

            try
            {
                var folders = await _graphClient.Me.MailFolders.GetAsync();
                if (folders?.Value != null)
                {
                    foreach (var folder in folders.Value)
                    {
                        string displayName = folder.DisplayName ?? "Folder";
                        // Translate folder names
                        if (displayName.Equals("Inbox", StringComparison.OrdinalIgnoreCase)) displayName = "📥 Hộp thư đến (Inbox)";
                        else if (displayName.Equals("Sent Items", StringComparison.OrdinalIgnoreCase)) displayName = "📤 Thư đã gửi (Sent)";
                        else if (displayName.Equals("Drafts", StringComparison.OrdinalIgnoreCase)) displayName = "📝 Thư nháp (Drafts)";
                        else if (displayName.Equals("Deleted Items", StringComparison.OrdinalIgnoreCase)) displayName = "🗑️ Thư đã xóa (Trash)";
                        else if (displayName.Equals("Junk Email", StringComparison.OrdinalIgnoreCase)) displayName = "🚫 Thư rác (Junk)";

                        foldersList.Add(new Tuple<string, string>(folder.Id!, displayName));
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to get mail folders: {ex.Message}");
            }

            return foldersList;
        }

        public static async Task<List<EmailMessage>> GetEmailsAsync(string folderId)
        {
            var emails = new List<EmailMessage>();
            if (_graphClient == null) return emails;

            try
            {
                var messages = await _graphClient.Me.MailFolders[folderId].Messages.GetAsync(config =>
                {
                    config.QueryParameters.Select = new[] { "id", "subject", "from", "toRecipients", "ccRecipients", "receivedDateTime", "hasAttachments", "body", "bodyPreview" };
                    config.QueryParameters.Top = 50; // Fetch latest 50 messages
                });

                if (messages?.Value != null)
                {
                    foreach (var msg in messages.Value)
                    {
                        var email = new EmailMessage
                        {
                            FilePath = $"o365://{msg.Id}",
                            Subject = string.IsNullOrWhiteSpace(msg.Subject) ? "(Không có tiêu đề)" : msg.Subject,
                            FromName = msg.From?.EmailAddress?.Name ?? "",
                            FromEmail = msg.From?.EmailAddress?.Address ?? "",
                            Date = msg.ReceivedDateTime?.DateTime.ToLocalTime(),
                            BodyText = msg.BodyPreview ?? ""
                        };

                        if (msg.ToRecipients != null && msg.ToRecipients.Any())
                        {
                            email.To = string.Join("; ", msg.ToRecipients.Select(r => r.EmailAddress?.Address ?? ""));
                        }

                        if (msg.CcRecipients != null && msg.CcRecipients.Any())
                        {
                            email.Cc = string.Join("; ", msg.CcRecipients.Select(r => r.EmailAddress?.Address ?? ""));
                        }

                        // Check if it has attachments
                        if (msg.HasAttachments == true)
                        {
                            // Place a dummy attachment to signal there are attachments
                            email.Attachments.Add(new EmailAttachment { FileName = "Đang tải tệp đính kèm..." });
                        }

                        emails.Add(email);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to get emails: {ex.Message}");
            }

            return emails;
        }

        public static async Task LoadEmailFullDetailsAsync(EmailMessage email)
        {
            if (_graphClient == null || !email.FilePath.StartsWith("o365://")) return;

            string messageId = email.FilePath.Substring("o365://".Length);

            try
            {
                var msg = await _graphClient.Me.Messages[messageId].GetAsync(config =>
                {
                    config.QueryParameters.Select = new[] { "body", "hasAttachments" };
                });

                if (msg != null)
                {
                    if (msg.Body?.ContentType == BodyType.Html)
                    {
                        email.BodyHtml = msg.Body.Content;
                        // Strip HTML tags for clean BodyText preview fallback
                        email.BodyText = System.Text.RegularExpressions.Regex.Replace(msg.Body.Content ?? "", "<.*?>", string.Empty);
                    }
                    else
                    {
                        email.BodyText = msg.Body?.Content ?? "";
                        email.BodyHtml = null;
                    }

                    // Load actual attachments
                    if (msg.HasAttachments == true)
                    {
                        email.Attachments.Clear();
                        var attachments = await _graphClient.Me.Messages[messageId].Attachments.GetAsync();
                        if (attachments?.Value != null)
                        {
                            foreach (var att in attachments.Value)
                            {
                                if (att is FileAttachment fileAtt)
                                {
                                    email.Attachments.Add(new EmailAttachment
                                    {
                                        FileName = fileAtt.Name ?? "unnamed",
                                        Data = fileAtt.ContentBytes ?? Array.Empty<byte>()
                                    });
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to load full email details: {ex.Message}");
            }
        }

        public static async Task<bool> SendEmailAsync(string to, string cc, string subject, string bodyHtml, List<string> attachmentPaths)
        {
            if (_graphClient == null) return false;

            try
            {
                var toRecipients = to.Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(email => new Recipient { EmailAddress = new EmailAddress { Address = email.Trim() } })
                    .ToList();

                var ccRecipients = string.IsNullOrWhiteSpace(cc) ? new List<Recipient>() :
                    cc.Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(email => new Recipient { EmailAddress = new EmailAddress { Address = email.Trim() } })
                    .ToList();

                var message = new Message
                {
                    Subject = subject,
                    Body = new ItemBody
                    {
                        ContentType = BodyType.Html,
                        Content = bodyHtml
                    },
                    ToRecipients = toRecipients,
                    CcRecipients = ccRecipients,
                    Attachments = new List<Attachment>()
                };

                foreach (var path in attachmentPaths)
                {
                    if (File.Exists(path))
                    {
                        byte[] fileBytes = await File.ReadAllBytesAsync(path);
                        message.Attachments.Add(new FileAttachment
                        {
                            Name = Path.GetFileName(path),
                            ContentBytes = fileBytes,
                            ContentType = "application/octet-stream"
                        });
                    }
                }

                var requestBody = new Microsoft.Graph.Me.SendMail.SendMailPostRequestBody
                {
                    Message = message,
                    SaveToSentItems = true
                };

                await _graphClient.Me.SendMail.PostAsync(requestBody);
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to send email: {ex.Message}");
                throw;
            }
        }
    }

    // Helper implementation for GraphServiceClient credential authentication
    public class TokenProviderCredential : Azure.Core.TokenCredential
    {
        private readonly Func<Task<string>> _tokenFactory;

        public TokenProviderCredential(Func<Task<string>> tokenFactory)
        {
            _tokenFactory = tokenFactory ?? throw new ArgumentNullException(nameof(tokenFactory));
        }

        public override Azure.Core.AccessToken GetToken(Azure.Core.TokenRequestContext requestContext, System.Threading.CancellationToken cancellationToken)
        {
            return GetTokenAsync(requestContext, cancellationToken).AsTask().GetAwaiter().GetResult();
        }

        public override async System.Threading.Tasks.ValueTask<Azure.Core.AccessToken> GetTokenAsync(Azure.Core.TokenRequestContext requestContext, System.Threading.CancellationToken cancellationToken)
        {
            string token = await _tokenFactory();
            return new Azure.Core.AccessToken(token, DateTimeOffset.UtcNow.AddHours(1));
        }
    }
}
