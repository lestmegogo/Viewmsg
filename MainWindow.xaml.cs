using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Collections.Generic;
using MsgViewer.Models;
using MsgViewer.Services;

namespace MsgViewer;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window
{
    private List<EmailMessage> _allEmails = new();
    private bool IsFocusedTabSelected = true;
    private EmailMessage? _currentEmail;

    public MainWindow()
    {
        InitializeComponent();
        UpdateEmailList();

        // Handle double-clicked email files from Windows Explorer
        var args = Environment.GetCommandLineArgs();
        if (args.Length > 1)
        {
            var filePath = args[1];
            if (File.Exists(filePath))
            {
                var ext = Path.GetExtension(filePath).ToLowerInvariant();
                if (ext == ".msg" || ext == ".eml")
                {
                    this.Loaded += (s, e) => {
                        try
                         {
                             var parsed = EmailParser.Parse(filePath);
                             if (!_allEmails.Any(em => em.FilePath == parsed.FilePath))
                             {
                                 _allEmails.Add(parsed);
                                 _allEmails = _allEmails.OrderByDescending(em => em.Date ?? DateTime.MinValue).ToList();
                                 UpdateEmailList();
                             }
                             
                             // Select and load the email details
                             var index = LstEmails.Items.Cast<EmailMessage>().ToList().FindIndex(em => em.FilePath == parsed.FilePath);
                             if (index >= 0)
                             {
                                 LstEmails.SelectedIndex = index;
                             }
                             else
                             {
                                 LoadEmailDetails(parsed);
                             }
                         }
                         catch (Exception ex)
                         {
                             MessageBox.Show($"Lỗi khi mở tệp email mặc định: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                         }
                    };
                }
            }
        }
    }



    private void BtnOpenFile_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new Microsoft.Win32.OpenFileDialog
        {
            Filter = "Email files (*.msg, *.eml)|*.msg;*.eml|Outlook Messages (*.msg)|*.msg|MIME Emails (*.eml)|*.eml|All files (*.*)|*.*",
            Title = "Chọn tệp email để mở"
        };
        if (dialog.ShowDialog() == true)
        {
            try
            {
                var parsed = EmailParser.Parse(dialog.FileName);
                if (!_allEmails.Any(em => em.FilePath == parsed.FilePath))
                {
                    _allEmails.Add(parsed);
                    _allEmails = _allEmails.OrderByDescending(em => em.Date ?? DateTime.MinValue).ToList();
                    UpdateEmailList();
                }
                
                // Select the opened email
                var index = LstEmails.Items.Cast<EmailMessage>().ToList().FindIndex(em => em.FilePath == parsed.FilePath);
                if (index >= 0)
                {
                    LstEmails.SelectedIndex = index;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi mở file email: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }

    private void BtnOpenFolder_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var dialog = new Microsoft.Win32.OpenFolderDialog
            {
                Title = "Chọn thư mục chứa các tệp email (.msg, .eml)"
            };
            if (dialog.ShowDialog() == true)
            {
                LoadFolder(dialog.FolderName);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Không thể mở hộp thoại chọn thư mục: {ex.Message}. Vui lòng kéo thả thư mục vào ứng dụng.", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }

    private void FocusedTab_Checked(object sender, RoutedEventArgs e)
    {
        IsFocusedTabSelected = true;
        UpdateEmailList();
    }

    private void OtherTab_Checked(object sender, RoutedEventArgs e)
    {
        IsFocusedTabSelected = false;
        UpdateEmailList();
    }

    private void UpdateEmailList()
    {
        if (_allEmails == null || LstEmails == null || BdrEmptyList == null) return;

        // Split emails: no-reply / news go to Other, rest to Focused
        var otherEmails = _allEmails.Where(e => 
            string.IsNullOrWhiteSpace(e.FromName) || 
            e.FromEmail.Contains("no-reply", StringComparison.OrdinalIgnoreCase) || 
            e.FromEmail.Contains("newsletter", StringComparison.OrdinalIgnoreCase) ||
            e.Subject.Contains("chúc mừng", StringComparison.OrdinalIgnoreCase) ||
            e.Subject.Contains("khuyến mại", StringComparison.OrdinalIgnoreCase)
        ).ToList();

        var focusedEmails = _allEmails.Except(otherEmails).ToList();

        if (IsFocusedTabSelected)
        {
            LstEmails.ItemsSource = focusedEmails;
            BdrEmptyList.Visibility = focusedEmails.Any() ? Visibility.Collapsed : Visibility.Visible;
        }
        else
        {
            LstEmails.ItemsSource = otherEmails;
            BdrEmptyList.Visibility = otherEmails.Any() ? Visibility.Collapsed : Visibility.Visible;
        }
    }

    private void LoadFolder(string folderPath)
    {
        try
        {
            TxtStatus.Text = $"Đang quét thư mục: {folderPath}...";
            var files = Directory.EnumerateFiles(folderPath, "*.*", SearchOption.AllDirectories)
                                 .Where(f => f.EndsWith(".msg", StringComparison.OrdinalIgnoreCase) || 
                                             f.EndsWith(".eml", StringComparison.OrdinalIgnoreCase))
                                 .ToList();

            if (files.Count == 0)
            {
                MessageBox.Show("Không tìm thấy tệp email .msg hoặc .eml nào trong thư mục này.", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Information);
                TxtStatus.Text = "Không tìm thấy email nào.";
                return;
            }

            int successCount = 0;
            foreach (var file in files)
            {
                try
                {
                    var parsed = EmailParser.Parse(file);
                    if (!_allEmails.Any(e => e.FilePath == parsed.FilePath))
                    {
                        _allEmails.Add(parsed);
                        successCount++;
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Failed to parse {file}: {ex.Message}");
                }
            }

            _allEmails = _allEmails.OrderByDescending(e => e.Date ?? DateTime.MinValue).ToList();
            
            UpdateEmailList();

            // Select the first item in the list if items exist
            if (LstEmails.Items.Count > 0)
            {
                LstEmails.SelectedIndex = 0;
            }
            
            TxtStatus.Text = $"Đã nạp thành công {successCount} email mới từ {folderPath}.";
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Lỗi khi đọc thư mục: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            TxtStatus.Text = "Lỗi khi nạp thư mục.";
        }
    }

    private void LstEmails_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (LstEmails.SelectedItem is EmailMessage selectedEmail)
        {
            LoadEmailDetails(selectedEmail);
        }
        else
        {
            GridPlaceholder.Visibility = Visibility.Visible;
        }
    }

    private void LoadEmailDetails(EmailMessage email)
    {
        _currentEmail = email;
        GridPlaceholder.Visibility = Visibility.Collapsed;
        GridDetail.Visibility = Visibility.Visible;

        TxtSubject.Text = string.IsNullOrWhiteSpace(email.Subject) ? "(Không có tiêu đề)" : email.Subject;
        TxtFrom.Text = email.FromDisplay;
        TxtTo.Text = "To: " + (string.IsNullOrWhiteSpace(email.To) ? "-" : email.To);

        if (string.IsNullOrWhiteSpace(email.Cc))
        {
            PanelCc.Visibility = Visibility.Collapsed;
        }
        else
        {
            PanelCc.Visibility = Visibility.Visible;
            TxtCc.Text = email.Cc;
        }

        TxtDate.Text = email.DateDisplay;
        TxtAvatar.Text = email.SenderInitials;

        // Safety banner shown by default
        BdrSafetyBanner.Visibility = Visibility.Visible;

        // Populate attachments
        if (email.Attachments.Count > 0)
        {
            CardAttachments.Visibility = Visibility.Visible;
            LstAttachments.ItemsSource = email.Attachments;
        }
        else
        {
            CardAttachments.Visibility = Visibility.Collapsed;
            LstAttachments.ItemsSource = null;
        }

        // Populate text view
        TxtBodyText.Text = string.IsNullOrWhiteSpace(email.BodyText) 
            ? "Không có nội dung văn bản thuần." 
            : email.BodyText;

        // Populate raw headers
        PopulateHeaders(email.FilePath);

        // Load content to WebBrowser safely
        try
        {
            SuppressScriptErrors(EmailWebBrowser, true);
            if (!string.IsNullOrWhiteSpace(email.BodyHtml))
            {
                EmailWebBrowser.NavigateToString(email.BodyHtml);
            }
            else
            {
                string plainTextEncoded = System.Net.WebUtility.HtmlEncode(email.BodyText ?? "");
                string simpleHtml = $"<html><body style='font-family:-apple-system,BlinkMacSystemFont,\"Segoe UI\",Roboto,Helvetica,Arial,sans-serif;white-space:pre-wrap;padding:20px;background-color:#ffffff;color:#1e293b;line-height:1.6;'>{plainTextEncoded}</body></html>";
                EmailWebBrowser.NavigateToString(simpleHtml);
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"WebBrowser navigation failed: {ex.Message}");
        }

        TxtStatus.Text = $"Đang hiển thị: {Path.GetFileName(email.FilePath)}";
    }

    private void PopulateHeaders(string filePath)
    {
        try
        {
            var ext = Path.GetExtension(filePath).ToLowerInvariant();
            if (ext == ".msg")
            {
                using var msg = new MsgReader.Outlook.Storage.Message(filePath);
                TxtHeaders.Text = string.IsNullOrWhiteSpace(msg.TransportMessageHeaders)
                    ? "Không tìm thấy thông tin Transport Message Headers trong file .msg này."
                    : msg.TransportMessageHeaders;
            }
            else if (ext == ".eml")
            {
                var eml = MsgReader.Mime.Message.Load(new FileInfo(filePath));
                var headers = eml.Headers.RawHeaders;
                if (headers != null && headers.Count > 0)
                {
                    var sb = new StringBuilder();
                    foreach (string? key in headers.AllKeys)
                    {
                        if (key != null)
                        {
                            var values = headers.GetValues(key);
                            if (values != null)
                            {
                                foreach (var val in values)
                                {
                                    sb.AppendLine($"{key}: {val}");
                                }
                            }
                        }
                    }
                    TxtHeaders.Text = sb.ToString();
                }
                else
                {
                    TxtHeaders.Text = "Không có thông tin headers.";
                }
            }
            else
            {
                TxtHeaders.Text = "Định dạng file không hỗ trợ đọc headers.";
            }
        }
        catch (Exception ex)
        {
            TxtHeaders.Text = $"Lỗi khi đọc headers: {ex.Message}";
        }
    }

    private void BtnSafetyBanner_Close(object sender, RoutedEventArgs e)
    {
        BdrSafetyBanner.Visibility = Visibility.Collapsed;
    }

    private void BtnSummarize_Click(object sender, RoutedEventArgs e)
    {
        if (_currentEmail == null) return;
        TxtSummaryContent.Text = GenerateSummary(_currentEmail);
        BdrSummaryPopup.Visibility = Visibility.Visible;
    }

    private void BtnCloseSummary_Click(object sender, RoutedEventArgs e)
    {
        BdrSummaryPopup.Visibility = Visibility.Collapsed;
    }

    private void BdrSummaryPopup_MouseDown(object sender, MouseButtonEventArgs e)
    {
        BdrSummaryPopup.Visibility = Visibility.Collapsed;
    }

    private void BdrSummaryPopup_InnerMouseDown(object sender, MouseButtonEventArgs e)
    {
        e.Handled = true;
    }

    private string GenerateSummary(EmailMessage email)
    {
        if (email == null) return "Không có email nào được chọn.";
        
        var sb = new StringBuilder();
        sb.AppendLine($"📌 TÓM TẮT EMAIL: {email.Subject}\n");
        sb.AppendLine($"• Người gửi: {email.FromName} ({email.FromEmail})");
        if (email.Date.HasValue)
        {
            sb.AppendLine($"• Thời gian gửi: {email.DateDisplay}");
        }
        
        if (!string.IsNullOrWhiteSpace(email.BodyText))
        {
            var lines = email.BodyText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                                      .Select(l => l.Trim())
                                      .Where(l => l.Length > 2)
                                      .ToList();

            var listItems = new List<string>();
            var keyPoints = new List<string>();

            foreach (var line in lines)
            {
                if (line.StartsWith("Dear", StringComparison.OrdinalIgnoreCase) || 
                    line.StartsWith("Kính gửi", StringComparison.OrdinalIgnoreCase) ||
                    line.StartsWith("Trân trọng", StringComparison.OrdinalIgnoreCase) ||
                    line.StartsWith("Thanks", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (line.StartsWith("-") || line.StartsWith("*") || line.StartsWith("•") || 
                    System.Text.RegularExpressions.Regex.IsMatch(line, @"^\d+[\.\s]"))
                {
                    listItems.Add(line);
                }
                else if (line.Contains("Cisco") || line.Contains("F5") || line.Contains("ISR") || 
                         line.Contains("Nexus") || line.Contains("Catalyst") || line.Contains("thiết bị") || 
                         line.Contains("yêu cầu") || line.Contains("tháo dỡ") || line.Contains("thời gian") ||
                         line.Contains("DC Duy Tân") || line.Contains("DC") || line.Contains("VETC"))
                {
                    keyPoints.Add(line);
                }
            }

            if (listItems.Any() || keyPoints.Any())
            {
                sb.AppendLine("\n📋 CÁC Ý CHÍNH & THIẾT BỊ ĐƯỢC ĐỀ CẬP:");
                var itemsToShow = listItems.Concat(keyPoints).Distinct().Take(12).ToList();
                foreach (var item in itemsToShow)
                {
                    var cleaned = System.Text.RegularExpressions.Regex.Replace(item, @"^[\-\*\•\s]+", "").Trim();
                    sb.AppendLine($"  - {cleaned}");
                }
            }
            else
            {
                sb.AppendLine("\n📝 TỔNG QUAN NỘI DUNG:");
                var overviewLines = lines.Where(l => !l.StartsWith("Dear", StringComparison.OrdinalIgnoreCase) && 
                                                     !l.StartsWith("Kính gửi", StringComparison.OrdinalIgnoreCase) &&
                                                     !l.Contains("Em gửi")).Take(4);
                foreach (var l in overviewLines)
                {
                    sb.AppendLine($"  - {l}");
                }
            }
        }
        else
        {
            sb.AppendLine("\n(Thư này không có nội dung văn bản)");
        }

        if (email.Attachments.Count > 0)
        {
            sb.AppendLine($"\n📎 TỆP ĐÍNH KÈM ({email.Attachments.Count} tệp):");
            foreach (var att in email.Attachments)
            {
                sb.AppendLine($"  - {att.FileName} ({att.SizeDisplay})");
            }
        }

        return sb.ToString();
    }

    private void BtnSaveAttachment_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button btn && btn.DataContext is EmailAttachment attachment)
        {
            var sfd = new Microsoft.Win32.SaveFileDialog
            {
                FileName = attachment.FileName,
                Title = "Lưu tệp đính kèm"
            };
            if (sfd.ShowDialog() == true)
            {
                try
                {
                    File.WriteAllBytes(sfd.FileName, attachment.Data);
                    MessageBox.Show("Đã lưu tệp đính kèm thành công!", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Không thể lưu tệp: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
    }

    private void BtnOpenAttachment_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button btn && btn.DataContext is EmailAttachment attachment)
        {
            try
            {
                string tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(tempDir);
                string tempFile = Path.Combine(tempDir, attachment.FileName);
                File.WriteAllBytes(tempFile, attachment.Data);

                var psi = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = tempFile,
                    UseShellExecute = true
                };
                System.Diagnostics.Process.Start(psi);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Không thể mở tệp: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }

    private void Window_DragEnter(object sender, DragEventArgs e)
    {
        if (e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            e.Effects = DragDropEffects.Copy;
        }
        else
        {
            e.Effects = DragDropEffects.None;
        }
    }

    private void Window_DragOver(object sender, DragEventArgs e)
    {
        if (e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            e.Effects = DragDropEffects.Copy;
            e.Handled = true;
        }
    }

    private void Window_Drop(object sender, DragEventArgs e)
    {
        if (e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            string[] paths = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (paths != null && paths.Length > 0)
            {
                int fileCount = 0;
                int folderCount = 0;

                foreach (var path in paths)
                {
                    if (Directory.Exists(path))
                    {
                        LoadFolder(path);
                        folderCount++;
                    }
                    else if (File.Exists(path))
                    {
                        var ext = Path.GetExtension(path).ToLowerInvariant();
                        if (ext == ".msg" || ext == ".eml")
                        {
                            try
                            {
                                var parsed = EmailParser.Parse(path);
                                if (!_allEmails.Any(em => em.FilePath == parsed.FilePath))
                                {
                                    _allEmails.Add(parsed);
                                    fileCount++;
                                }
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine($"Error parsing drag-dropped file {path}: {ex.Message}");
                            }
                        }
                    }
                }

                if (fileCount > 0)
                {
                    _allEmails = _allEmails.OrderByDescending(em => em.Date ?? DateTime.MinValue).ToList();
                    UpdateEmailList();
                    TxtStatus.Text = $"Đã nạp thêm {fileCount} tệp email.";
                }
            }
        }
    }

    private void SuppressScriptErrors(WebBrowser browser, bool silent)
    {
        try
        {
            var activeX = browser.GetType().GetProperty("ActiveXInstance", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)?.GetValue(browser);
            if (activeX != null)
            {
                activeX.GetType().InvokeMember("Silent", System.Reflection.BindingFlags.SetProperty, null, activeX, new object[] { silent });
            }
        }
        catch
        {
            // Ignore reflection errors
        }
    }
}