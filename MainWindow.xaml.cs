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
using XstReader;
using XstReader.ElementProperties;


namespace MsgViewer;

/// <summary>
/// View model class representing a node in the Outlook Folder TreeView
/// </summary>
public class PstFolderNode
{
    public string Name { get; set; } = "";
    public string Icon { get; set; } = "📁";
    public XstFolder? Folder { get; set; } = null;
    public string Id { get; set; } = "";
    public bool IsOffice365 { get; set; } = false;
    public List<PstFolderNode> SubFolders { get; } = new();
    public int UnreadCount { get; set; } = 0;
    public string UnreadDisplay => UnreadCount > 0 ? $" ({UnreadCount})" : "";
    public bool IsSelected { get; set; } = false;
}

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window
{
    private List<EmailMessage> _allEmails = new();
    private bool IsFocusedTabSelected = true;
    private EmailMessage? _currentEmail;
    private XstFile? _currentPstFile;
    private string _currentSearchText = "";

    public MainWindow()
    {
        InitializeComponent();
        InitializeWebView();
        InitializeOffice365();
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
                else if (ext == ".pst" || ext == ".ost")
                {
                    this.Loaded += (s, e) => {
                        LoadPstFile(filePath);
                    };
                }
            }
        }
    }




    private void ClosePstView()
    {
        ColFolderTree.Width = new GridLength(0);
        ColFolderTree.MinWidth = 0;
        if (Splitter1 != null) Splitter1.Visibility = Visibility.Collapsed;
        PstFolderPanel.Visibility = Visibility.Collapsed;
        _currentPstFile?.Dispose();
        _currentPstFile = null;
    }

    private async void InitializeWebView()
    {
        try
        {
            await EmailWebView.EnsureCoreWebView2Async(null);
        }
        catch (Exception)
        {
            MessageBoxResult result = MessageBox.Show(
                "Không thể khởi chạy trình duyệt Edge Chromium (WebView2). Có thể máy tính của bạn chưa được cài đặt WebView2 Runtime.\n\nBạn có muốn tải xuống và cài đặt WebView2 Runtime từ trang chủ của Microsoft ngay bây giờ không?",
                "Yêu cầu thành phần hệ thống",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (result == MessageBoxResult.Yes)
            {
                try
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = "https://go.microsoft.com/fwlink/p/?LinkId=2124703",
                        UseShellExecute = true
                    });
                }
                catch (Exception openEx)
                {
                    MessageBox.Show($"Không thể mở trình duyệt để tải xuống: {openEx.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
    }

    private async void InitializeOffice365()
    {
        try
        {
            await Office365Service.InitializeAsync();
            UpdateO365UIState();
            if (Office365Service.IsSignedIn)
            {
                await LoadO365FoldersAsync();
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"O365 silent init failed: {ex.Message}");
        }
    }

    private void UpdateO365UIState()
    {
        if (Office365Service.IsSignedIn)
        {
            GridO365SignedOut.Visibility = Visibility.Collapsed;
            GridO365SignedIn.Visibility = Visibility.Visible;
            TxtO365Name.Text = Office365Service.UserDisplayName ?? "User";
            TxtO365Email.Text = Office365Service.UserEmail ?? "";
            
            if (!string.IsNullOrWhiteSpace(Office365Service.UserDisplayName))
            {
                var parts = Office365Service.UserDisplayName.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                TxtO365Avatar.Text = parts.Length >= 2 
                    ? (parts[0][0].ToString() + parts[parts.Length - 1][0].ToString()).ToUpper()
                    : Office365Service.UserDisplayName.Substring(0, Math.Min(2, Office365Service.UserDisplayName.Length)).ToUpper();
            }
            else
            {
                TxtO365Avatar.Text = "U";
            }
        }
        else
        {
            GridO365SignedOut.Visibility = Visibility.Visible;
            GridO365SignedIn.Visibility = Visibility.Collapsed;
        }
    }

    private async Task LoadO365FoldersAsync()
    {
        try
        {
            var folders = await Office365Service.GetFoldersAsync();
            
            // Show FolderTree column and splitter
            ColFolderTree.Width = new GridLength(240);
            ColFolderTree.MinWidth = 150;
            if (Splitter1 != null) Splitter1.Visibility = Visibility.Visible;
            PstFolderPanel.Visibility = Visibility.Visible;
            
            // Clear TreeView and add Office 365 folders
            TreePstFolders.ItemsSource = null;
            
            var rootNodes = folders.Select(f => new PstFolderNode
            {
                Name = f.Item2,
                Id = f.Item1,
                IsOffice365 = true,
                Icon = f.Item2.Contains("Inbox") ? "📥" : 
                       f.Item2.Contains("Sent") ? "📤" : 
                       f.Item2.Contains("Drafts") ? "📝" : 
                       f.Item2.Contains("Trash") ? "🗑️" : "📁"
            }).ToList();
            
            TreePstFolders.ItemsSource = rootNodes;
            TxtStatus.Text = "Đã đồng bộ danh mục thư Office 365.";

            var inboxNode = rootNodes.FirstOrDefault(n => n.Name.Contains("Inbox") || n.Name.Contains("Hộp thư đến"));
            if (inboxNode != null)
            {
                inboxNode.IsSelected = true;
                TreePstFolders_SelectedItemChanged(TreePstFolders, new RoutedPropertyChangedEventArgs<object>(inboxNode, inboxNode));
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Lỗi tải thư mục Office 365: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async void BtnO365SignIn_Click(object sender, RoutedEventArgs e)
    {
        TxtStatus.Text = "Đang đăng nhập Office 365...";
        bool success = await Office365Service.SignInAsync();
        if (success)
        {
            UpdateO365UIState();
            await LoadO365FoldersAsync();
            TxtStatus.Text = "Đăng nhập thành công!";
        }
        else
        {
            TxtStatus.Text = "Đăng nhập thất bại hoặc bị hủy.";
            MessageBox.Show("Đăng nhập Office 365 thất bại. Vui lòng kiểm tra lại kết nối và tài khoản.", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private async void BtnO365SignOut_Click(object sender, RoutedEventArgs e)
    {
        if (MessageBox.Show("Bạn có muốn đăng xuất tài khoản Office 365 không?", "Đăng xuất", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
        {
            await Office365Service.SignOutAsync();
            UpdateO365UIState();
            ClosePstView();
            _allEmails.Clear();
            UpdateEmailList();
            TxtStatus.Text = "Đã đăng xuất tài khoản.";
        }
    }

    private void BtnO365Compose_Click(object sender, RoutedEventArgs e)
    {
        var composeWin = new ComposeWindow { Owner = this };
        composeWin.ShowDialog();
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
                ClosePstView();
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
                ClosePstView();
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

        // Apply Search Filtering first
        var keyword = _currentSearchText?.Trim();
        var emailsToFilter = _allEmails.AsEnumerable();

        if (!string.IsNullOrWhiteSpace(keyword))
        {
            int scopeIndex = CboSearchScope?.SelectedIndex ?? 0;
            emailsToFilter = emailsToFilter.Where(e => {
                switch (scopeIndex)
                {
                    case 1: // Subject only
                        return e.Subject.Contains(keyword, StringComparison.OrdinalIgnoreCase);
                    case 2: // Sender only
                        return e.FromName.Contains(keyword, StringComparison.OrdinalIgnoreCase) ||
                               e.FromEmail.Contains(keyword, StringComparison.OrdinalIgnoreCase);
                    case 3: // Recipient only
                        return e.To.Contains(keyword, StringComparison.OrdinalIgnoreCase);
                    case 0: // All fields
                    default:
                        return e.Subject.Contains(keyword, StringComparison.OrdinalIgnoreCase) ||
                               e.FromName.Contains(keyword, StringComparison.OrdinalIgnoreCase) ||
                               e.FromEmail.Contains(keyword, StringComparison.OrdinalIgnoreCase) ||
                               e.To.Contains(keyword, StringComparison.OrdinalIgnoreCase) ||
                               (e.BodyText != null && e.BodyText.Contains(keyword, StringComparison.OrdinalIgnoreCase));
                }
            });
        }

        // Apply Date Filtering
        if (CboDateFilter != null && CboDateFilter.SelectedIndex > 0)
        {
            emailsToFilter = emailsToFilter.Where(e => {
                if (!e.Date.HasValue) return false;
                var date = e.Date.Value.Date;
                switch (CboDateFilter.SelectedIndex)
                {
                    case 1: // Today
                        return date == DateTime.Today;
                    case 2: // Yesterday
                        return date == DateTime.Today.AddDays(-1);
                    case 3: // This week
                        int diff = (7 + (DateTime.Today.DayOfWeek - DayOfWeek.Monday)) % 7;
                        DateTime startOfWeek = DateTime.Today.AddDays(-1 * diff).Date;
                        return date >= startOfWeek;
                    case 4: // This month
                        var startOfMonth = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
                        return date >= startOfMonth;
                    case 5: // Last month
                        var startOfLastMonth = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1).AddMonths(-1);
                        var startOfThisMonth = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
                        return date >= startOfLastMonth && date < startOfThisMonth;
                    case 6: // Custom range
                        var start = DpStartDate?.SelectedDate;
                        var end = DpEndDate?.SelectedDate;
                        if (start.HasValue && end.HasValue)
                        {
                            return date >= start.Value.Date && date <= end.Value.Date;
                        }
                        else if (start.HasValue)
                        {
                            return date >= start.Value.Date;
                        }
                        else if (end.HasValue)
                        {
                            return date <= end.Value.Date;
                        }
                        return true;
                    default:
                        return true;
                }
            });
        }

        // Apply Attachment Filtering
        if (ChkHasAttachments != null && ChkHasAttachments.IsChecked == true)
        {
            emailsToFilter = emailsToFilter.Where(e => e.HasAttachments);
        }

        var filteredList = emailsToFilter.ToList();

        // Split emails: no-reply / news go to Other, rest to Focused
        var otherEmails = filteredList.Where(e => 
            string.IsNullOrWhiteSpace(e.FromName) || 
            e.FromEmail.Contains("no-reply", StringComparison.OrdinalIgnoreCase) || 
            e.FromEmail.Contains("newsletter", StringComparison.OrdinalIgnoreCase) ||
            e.Subject.Contains("chúc mừng", StringComparison.OrdinalIgnoreCase) ||
            e.Subject.Contains("khuyến mại", StringComparison.OrdinalIgnoreCase)
        ).ToList();

        var focusedEmails = filteredList.Except(otherEmails).ToList();

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
            GridDetail.Visibility = Visibility.Collapsed;
        }
    }


    private async void LoadEmailDetails(EmailMessage email)
    {
        _currentEmail = email;
        GridPlaceholder.Visibility = Visibility.Collapsed;
        GridDetail.Visibility = Visibility.Visible;

        // Nạp lazy body & attachments từ XstMessage nếu là thư trong PST
        if (email.RawXstMessage != null)
        {
            PstParser.LoadMessageDetails(email);
        }
        else if (email.FilePath.StartsWith("o365://"))
        {
            TxtStatus.Text = "Đang tải chi tiết thư từ Office 365...";
            await Office365Service.LoadEmailFullDetailsAsync(email);
        }

        TxtSubject.Text = string.IsNullOrWhiteSpace(email.Subject) ? "(Không có tiêu đề)" : email.Subject;
        TxtFrom.Text = email.FromDisplay;
        TxtTo.Text = "To: " + (string.IsNullOrWhiteSpace(email.To) ? "-" : email.To);

        if (string.IsNullOrWhiteSpace(email.Cc))
        {
            TxtCc.Visibility = Visibility.Collapsed;
        }
        else
        {
            TxtCc.Visibility = Visibility.Visible;
            TxtCc.Text = "Cc: " + email.Cc;
        }


        TxtDate.Text = email.DateDisplay;
        TxtAvatar.Text = email.SenderInitials;

        // Safety banner đã loại bỏ

        // Nạp đính kèm
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

        // Nạp text thuần
        TxtBodyText.Text = string.IsNullOrWhiteSpace(email.BodyText) 
            ? "Không có nội dung văn bản thuần." 
            : email.BodyText;

        // Nạp SMTP headers
        PopulateHeaders(email);

        // Hiển thị WebView2 an toàn
        try
        {
            if (EmailWebView.CoreWebView2 == null)
            {
                await EmailWebView.EnsureCoreWebView2Async(null);
            }

            if (EmailWebView.CoreWebView2 != null)
            {
                if (!string.IsNullOrWhiteSpace(email.BodyHtml))
                {
                    string html = email.BodyHtml;
                    string cssStyle = "<style>body { padding: 20px !important; }</style>";
                    if (html.Contains("<head>", StringComparison.OrdinalIgnoreCase))
                    {
                        int headIndex = html.IndexOf("<head>", StringComparison.OrdinalIgnoreCase);
                        html = html.Insert(headIndex + 6, cssStyle);
                    }
                    else if (html.Contains("<html>", StringComparison.OrdinalIgnoreCase))
                    {
                        int htmlIndex = html.IndexOf("<html>", StringComparison.OrdinalIgnoreCase);
                        html = html.Insert(htmlIndex + 6, "<head>" + cssStyle + "</head>");
                    }
                    else
                    {
                        html = cssStyle + html;
                    }
                    EmailWebView.NavigateToString(html);
                }
                else
                {
                    string plainTextEncoded = System.Net.WebUtility.HtmlEncode(email.BodyText ?? "");
                    string simpleHtml = $"<html><body style='font-family:-apple-system,BlinkMacSystemFont,\"Segoe UI\",Roboto,Helvetica,Arial,sans-serif;white-space:pre-wrap;padding:20px;background-color:#ffffff;color:#1e293b;line-height:1.6;'>{plainTextEncoded}</body></html>";
                    EmailWebView.NavigateToString(simpleHtml);
                }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"WebView2 navigation failed: {ex.Message}");
        }

        if (email.RawXstMessage != null)
        {
            TxtStatus.Text = $"Đang hiển thị thư trong PST: {email.Subject}";
        }
        else
        {
            TxtStatus.Text = $"Đang hiển thị: {Path.GetFileName(email.FilePath)}";
        }
    }

    private void PopulateHeaders(EmailMessage email)
    {
        try
        {
            if (email.RawXstMessage is XstMessage xstMsg)
            {
                var headersProp = xstMsg.Properties[PropertyCanonicalName.PidTagTransportMessageHeaders];
                TxtHeaders.Text = headersProp?.Value?.ToString() ?? "Không tìm thấy thông tin Transport Message Headers trong thư PST này.";
                return;
            }

            var filePath = email.FilePath;
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


    // Warning banner click handler removed

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
                        else if (ext == ".pst" || ext == ".ost")
                        {
                            LoadPstFile(path);
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



    // ----- CÁC PHƯƠNG THỨC HỖ TRỢ ĐỌC FILE PST/OST & TÌM KIẾM -----

    private void BtnOpenPst_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new Microsoft.Win32.OpenFileDialog
        {
            Filter = "Outlook Data Files (*.pst, *.ost)|*.pst;*.ost|Personal Storage Table (*.pst)|*.pst|Offline Storage Table (*.ost)|*.ost|All files (*.*)|*.*",
            Title = "Chọn tệp dữ liệu Outlook để mở"
        };
        if (dialog.ShowDialog() == true)
        {
            LoadPstFile(dialog.FileName);
        }
    }

    private void LoadPstFile(string filePath)
    {
        try
        {
            TxtStatus.Text = $"Đang đọc tệp Outlook: {filePath}...";
            
            // Giải phóng file cũ nếu đang mở
            ClosePstView();

            // Mở file PST/OST mới
            _currentPstFile = new XstFile(filePath);

            // Cấu hình giao diện chế độ 3 cột
            ColFolderTree.Width = new GridLength(240);
            ColFolderTree.MinWidth = 150;
            Splitter1.Visibility = Visibility.Visible;
            PstFolderPanel.Visibility = Visibility.Visible;

            // Xây dựng cây thư mục đệ quy
            var rootNodes = new List<PstFolderNode>();
            var rootNode = CreateFolderNode(_currentPstFile.RootFolder);
            rootNodes.Add(rootNode);
            TreePstFolders.ItemsSource = rootNodes;

            // Tự động mở rộng node gốc
            if (TreePstFolders.ItemContainerGenerator.ContainerFromItem(rootNode) is TreeViewItem tvi)
            {
                tvi.IsExpanded = true;
            }

            TxtStatus.Text = $"Đã mở thành công tệp dữ liệu Outlook: {Path.GetFileName(filePath)}";
            _allEmails.Clear();
            UpdateEmailList();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Lỗi khi mở tệp PST: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            TxtStatus.Text = "Lỗi khi nạp tệp PST.";
            ClosePstView();
        }
    }

    private PstFolderNode CreateFolderNode(XstFolder folder)
    {
        var node = new PstFolderNode
        {
            Name = folder.DisplayName ?? "(Thư mục không tên)",
            Folder = folder,
            UnreadCount = folder.ContentUnreadCount
        };

        // Ánh xạ biểu tượng icon thư mục Outlook
        var nameLower = (folder.DisplayName ?? "").ToLowerInvariant();
        if (nameLower.Contains("inbox") || nameLower.Contains("hộp thư đến")) node.Icon = "📥";
        else if (nameLower.Contains("sent") || nameLower.Contains("thư đã gửi")) node.Icon = "📤";
        else if (nameLower.Contains("draft") || nameLower.Contains("thư nháp")) node.Icon = "📝";
        else if (nameLower.Contains("delete") || nameLower.Contains("thùng rác") || nameLower.Contains("trash")) node.Icon = "🗑️";
        else if (nameLower.Contains("junk") || nameLower.Contains("thư rác") || nameLower.Contains("spam")) node.Icon = "🚫";
        else if (nameLower.Contains("archive") || nameLower.Contains("lưu trữ")) node.Icon = "📦";
        else if (nameLower.Contains("calendar") || nameLower.Contains("lịch")) node.Icon = "📅";
        else if (nameLower.Contains("contact") || nameLower.Contains("danh bạ")) node.Icon = "👥";
        else node.Icon = "📁";

        if (folder.Folders != null)
        {
            foreach (var sub in folder.Folders)
            {
                node.SubFolders.Add(CreateFolderNode(sub));
            }
        }
        return node;
    }

    private async void TreePstFolders_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
    {
        if (e.NewValue is PstFolderNode node)
        {
            try
            {
                TxtStatus.Text = $"Đang đọc danh sách thư từ thư mục: {node.Name}...";
                _allEmails.Clear();

                if (node.IsOffice365)
                {
                    var o365Emails = await Office365Service.GetEmailsAsync(node.Id);
                    _allEmails.AddRange(o365Emails);
                }
                else if (node.Folder != null)
                {
                    var messages = node.Folder.Messages ?? node.Folder.GetMessages();
                    if (messages != null)
                    {
                        foreach (var msg in messages)
                        {
                            // Map tóm tắt thư (để load danh sách cực nhanh không tốn RAM)
                            _allEmails.Add(PstParser.MapMessageSummary(msg, _currentPstFile?.Path ?? ""));
                        }
                    }
                }

                _allEmails = _allEmails.OrderByDescending(em => em.Date ?? DateTime.MinValue).ToList();
                UpdateEmailList();
                TxtStatus.Text = $"Đã tải {_allEmails.Count} email trong thư mục: {node.Name}";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi tải email từ thư mục: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                TxtStatus.Text = "Lỗi khi nạp email.";
            }
        }
    }

    // Xử lý sự kiện Tìm kiếm thời gian thực
    private void TxtSearch_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (TxtSearch == null || TxtSearchPlaceholder == null || BtnClearSearch == null) return;

        _currentSearchText = TxtSearch.Text;
        TxtSearchPlaceholder.Visibility = string.IsNullOrEmpty(_currentSearchText) ? Visibility.Visible : Visibility.Collapsed;
        BtnClearSearch.Visibility = string.IsNullOrEmpty(_currentSearchText) ? Visibility.Collapsed : Visibility.Visible;
        UpdateEmailList();
    }

    private void TxtSearch_GotFocus(object sender, RoutedEventArgs e)
    {
        TxtSearchPlaceholder.Visibility = Visibility.Collapsed;
    }

    private void TxtSearch_LostFocus(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(TxtSearch.Text))
        {
            TxtSearchPlaceholder.Visibility = Visibility.Visible;
        }
    }

    private void BtnClearSearch_Click(object sender, RoutedEventArgs e)
    {
        TxtSearch.Text = "";
    }

    private void BtnToggleFilters_Checked(object sender, RoutedEventArgs e)
    {
        if (PanelAdvancedFilters != null)
        {
            PanelAdvancedFilters.Visibility = Visibility.Visible;
        }
    }

    private void BtnToggleFilters_Unchecked(object sender, RoutedEventArgs e)
    {
        if (PanelAdvancedFilters != null)
        {
            PanelAdvancedFilters.Visibility = Visibility.Collapsed;
            if (CboSearchScope != null) CboSearchScope.SelectedIndex = 0;
            if (CboDateFilter != null) CboDateFilter.SelectedIndex = 0;
            if (ChkHasAttachments != null) ChkHasAttachments.IsChecked = false;
            if (DpStartDate != null) DpStartDate.SelectedDate = null;
            if (DpEndDate != null) DpEndDate.SelectedDate = null;
            UpdateEmailList();
        }
    }

    private void FilterOption_Changed(object sender, RoutedEventArgs e)
    {
        UpdateEmailList();
    }

    private void FilterOption_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        UpdateEmailList();
    }

    private void CboDateFilter_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (CboDateFilter == null || GridCustomDateRange == null) return;

        if (CboDateFilter.SelectedIndex == 6) // "Tự chọn ngày..."
        {
            GridCustomDateRange.Visibility = Visibility.Visible;
        }
        else
        {
            GridCustomDateRange.Visibility = Visibility.Collapsed;
            if (DpStartDate != null) DpStartDate.SelectedDate = null;
            if (DpEndDate != null) DpEndDate.SelectedDate = null;
        }

        UpdateEmailList();
    }

    private void DatePicker_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
    {
        UpdateEmailList();
    }

    protected override void OnClosed(EventArgs e)
    {
        // Giải phóng file PST khi đóng ứng dụng
        _currentPstFile?.Dispose();
        base.OnClosed(e);
    }
}