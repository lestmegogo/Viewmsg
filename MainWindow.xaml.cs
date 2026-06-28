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
using System.Windows.Documents;
using System.Windows.Media;


namespace MsgViewer;

/// <summary>
/// View model class representing a node in the Outlook Folder TreeView
/// </summary>
public class PstFolderNode : System.ComponentModel.INotifyPropertyChanged
{
    private string _name = "";
    public string Name 
    { 
        get => _name; 
        set { _name = value; OnPropertyChanged(nameof(Name)); } 
    }

    private string _icon = "📁";
    public string Icon 
    { 
        get => _icon; 
        set { _icon = value; OnPropertyChanged(nameof(Icon)); } 
    }

    public XstFolder? Folder { get; set; } = null;
    public string Id { get; set; } = "";
    public string PstFilePath { get; set; } = "";
    public bool IsOffice365 { get; set; } = false;
    public List<PstFolderNode> SubFolders { get; } = new();

    private int _unreadCount = 0;
    public int UnreadCount
    {
        get => _unreadCount;
        set
        {
            _unreadCount = value;
            OnPropertyChanged(nameof(UnreadCount));
            OnPropertyChanged(nameof(UnreadDisplay));
            OnPropertyChanged(nameof(UnreadCountText));
        }
    }

    public string UnreadDisplay => UnreadCount > 0 ? $" ({UnreadCount})" : "";
    public string UnreadCountText => UnreadCount > 0 ? UnreadCount.ToString() : "";
    public bool IsSelected { get; set; } = false;

    public event System.ComponentModel.PropertyChangedEventHandler? PropertyChanged;
    protected void OnPropertyChanged(string propName)
    {
        PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(propName));
    }
}

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window
{
    private List<EmailMessage> _allEmails = new();
    private bool _showUnreadOnlyTab = false;
    private EmailMessage? _currentEmail;
    private Dictionary<string, XstFile> _activePstFiles = new();
    private string _currentSearchText = "";
    private System.Windows.Threading.DispatcherTimer? _o365NewMailTimer;
    private string? _activeFolderId;
    private System.Windows.Threading.DispatcherTimer? _owaEmailGrabberTimer;
    private System.Windows.Threading.DispatcherTimer? _searchDebounceTimer;

    private string _inlineActionType = "";
    private List<string> _inlineAttachments = new();
    private string? _inlineDraftId;
    private System.Windows.Threading.DispatcherTimer? _inlineDraftTimer;

    private string _currentSortTag = "DateDesc";

    public static readonly DependencyProperty SearchKeywordProperty =
        DependencyProperty.Register("SearchKeyword", typeof(string), typeof(MainWindow), new PropertyMetadata(""));

    public string SearchKeyword
    {
        get => (string)GetValue(SearchKeywordProperty);
        set => SetValue(SearchKeywordProperty, value);
    }

    public MainWindow()
    {
        InitializeComponent();
        InitializeWebView();
        InitializeOffice365();

        // Restore saved PST files session
        LoadPstSessions();
        if (_activePstFiles.Count > 0)
        {
            ColFolderTree.Width = new GridLength(240);
            ColFolderTree.MinWidth = 150;
            if (Splitter1 != null) Splitter1.Visibility = Visibility.Visible;
            PstFolderPanel.Visibility = Visibility.Visible;
            RefreshFolderTreeView();
        }

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




    private static readonly string PstSessionFilePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "MsgViewer",
        "pst_sessions.json");

    private void SavePstSessions()
    {
        try
        {
            string dir = Path.GetDirectoryName(PstSessionFilePath)!;
            if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);

            var paths = new List<string>(_activePstFiles.Keys);
            string json = System.Text.Json.JsonSerializer.Serialize(paths);
            File.WriteAllText(PstSessionFilePath, json);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Failed to save PST sessions: {ex.Message}");
        }
    }

    private void LoadPstSessions()
    {
        try
        {
            if (File.Exists(PstSessionFilePath))
            {
                string json = File.ReadAllText(PstSessionFilePath);
                var paths = System.Text.Json.JsonSerializer.Deserialize<List<string>>(json);
                if (paths != null)
                {
                    foreach (var path in paths)
                    {
                        if (File.Exists(path) && !_activePstFiles.ContainsKey(path))
                        {
                            try
                            {
                                var pstFile = new XstFile(path);
                                _activePstFiles[path] = pstFile;
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine($"Failed to restore PST {path}: {ex.Message}");
                            }
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Failed to load PST sessions: {ex.Message}");
        }
    }

    private void RefreshFolderTreeView()
    {
        var rootNodes = new List<PstFolderNode>();



        // 2. Add each active PST file root node
        foreach (var kvp in _activePstFiles)
        {
            try
            {
                var pstFile = kvp.Value;
                var pstRootNode = CreateFolderNode(pstFile.RootFolder, kvp.Key);
                pstRootNode.Name = Path.GetFileName(kvp.Key);
                pstRootNode.Id = kvp.Key; // File path serves as ID
                pstRootNode.IsOffice365 = false;
                pstRootNode.Icon = "📦"; // Archive/PST data file icon
                
                rootNodes.Add(pstRootNode);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to build tree node for PST {kvp.Key}: {ex.Message}");
            }
        }

        TreePstFolders.ItemsSource = null;
        TreePstFolders.ItemsSource = rootNodes;
    }

    private void ClosePstView()
    {
        foreach (var pst in _activePstFiles.Values)
        {
            pst.Dispose();
        }
        _activePstFiles.Clear();
        SavePstSessions();
        RefreshFolderTreeView();

        if (!IsWebSessionActive())
        {
            ColFolderTree.Width = new GridLength(0);
            ColFolderTree.MinWidth = 0;
            if (Splitter1 != null) Splitter1.Visibility = Visibility.Collapsed;
            PstFolderPanel.Visibility = Visibility.Collapsed;
        }
    }

    private PstFolderNode? GetNodeFromSender(object sender)
    {
        if (sender is MenuItem menuItem)
        {
            DependencyObject current = menuItem;
            while (current != null && !(current is ContextMenu))
            {
                current = LogicalTreeHelper.GetParent(current) ?? VisualTreeHelper.GetParent(current);
            }
            if (current is ContextMenu menu)
            {
                return (menu.DataContext as PstFolderNode) ?? ((menu.PlacementTarget as FrameworkElement)?.DataContext as PstFolderNode);
            }
        }
        return null;
    }

    private void CtxFolderMenu_Opened(object sender, RoutedEventArgs e)
    {
        if (sender is ContextMenu menu)
        {
            var node = (menu.DataContext as PstFolderNode) ?? ((menu.PlacementTarget as FrameworkElement)?.DataContext as PstFolderNode);
            if (node == null) return;

            MenuItem? statsItem = null;
            MenuItem? closePstItem = null;
            Separator? sepClosePst = null;

            // Scan items collection dynamically to find components
            foreach (var item in menu.Items)
            {
                if (item is MenuItem mi)
                {
                    string header = mi.Header?.ToString() ?? "";
                    if (header.Contains("items") || header.Contains("stats") || !mi.IsEnabled)
                    {
                        statsItem = mi;
                    }
                    else if (header.Contains("Đóng tệp dữ liệu"))
                    {
                        closePstItem = mi;
                    }
                }
                else if (item is Separator sep)
                {
                    // Track separator, the last separator corresponds to Close PST separator
                    sepClosePst = sep;
                }
            }

            // Update stats header
            if (statsItem != null)
            {
                int totalItems = 0;
                int unreadItems = node.UnreadCount;

                if (node.IsOffice365)
                {
                    var cached = OfflineCacheService.LoadEmails(node.Id);
                    totalItems = cached?.Count ?? 0;
                }
                else if (node.Folder != null)
                {
                    totalItems = node.Folder.ContentCount;
                    unreadItems = node.Folder.ContentUnreadCount;
                }

                statsItem.Header = $"{totalItems:N0} items ({unreadItems:N0} unread)";
            }

            // Control visibility of Close PST option
            bool isPstRoot = !node.IsOffice365 && _activePstFiles.ContainsKey(node.Id);
            if (closePstItem != null) closePstItem.Visibility = isPstRoot ? Visibility.Visible : Visibility.Collapsed;
            if (sepClosePst != null) sepClosePst.Visibility = isPstRoot ? Visibility.Visible : Visibility.Collapsed;
        }
    }

    private async void MnuCreateSubfolder_Click(object sender, RoutedEventArgs e)
    {
        var node = GetNodeFromSender(sender);
        if (node != null)
        {
            if (!node.IsOffice365)
            {
                MessageBox.Show("Thư mục từ tệp dữ liệu PST/OST cục bộ ở chế độ chỉ đọc.", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            string newName = Microsoft.VisualBasic.Interaction.InputBox("Nhập tên thư mục con mới:", "Tạo thư mục con", "");
            if (!string.IsNullOrWhiteSpace(newName))
            {
                TxtStatus.Text = $"Đang tạo thư mục: {newName}...";
                bool success = await Office365Service.CreateFolderAsync(node.Id, newName);
                if (success)
                {
                    await LoadO365FoldersAsync();
                    TxtStatus.Text = $"Đã tạo thành công thư mục con: {newName}";
                }
                else
                {
                    MessageBox.Show("Tạo thư mục con thất bại. Vui lòng kiểm tra lại quyền truy cập.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
    }

    private async void MnuRenameFolder_Click(object sender, RoutedEventArgs e)
    {
        var node = GetNodeFromSender(sender);
        if (node != null)
        {
            if (!node.IsOffice365)
            {
                MessageBox.Show("Thư mục từ tệp dữ liệu PST/OST cục bộ ở chế độ chỉ đọc.", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            string newName = Microsoft.VisualBasic.Interaction.InputBox("Nhập tên mới cho thư mục:", "Đổi tên thư mục", node.Name);
            if (!string.IsNullOrWhiteSpace(newName) && newName != node.Name)
            {
                TxtStatus.Text = $"Đang đổi tên thư mục thành: {newName}...";
                bool success = await Office365Service.RenameFolderAsync(node.Id, newName);
                if (success)
                {
                    await LoadO365FoldersAsync();
                    TxtStatus.Text = $"Đã đổi tên thư mục thành: {newName}";
                }
                else
                {
                    MessageBox.Show("Đổi tên thư mục thất bại.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
    }

    private async void MnuDeleteFolder_Click(object sender, RoutedEventArgs e)
    {
        var node = GetNodeFromSender(sender);
        if (node != null)
        {
            if (!node.IsOffice365)
            {
                MessageBox.Show("Thư mục từ tệp dữ liệu PST/OST cục bộ ở chế độ chỉ đọc.", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (MessageBox.Show($"Bạn có chắc chắn muốn xóa thư mục '{node.Name}' và toàn bộ email bên trong không?", "Xóa thư mục", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
            {
                TxtStatus.Text = $"Đang xóa thư mục: {node.Name}...";
                bool success = await Office365Service.DeleteFolderAsync(node.Id);
                if (success)
                {
                    await LoadO365FoldersAsync();
                    _allEmails.Clear();
                    UpdateEmailList();
                    TxtStatus.Text = $"Đã xóa thư mục: {node.Name}";
                }
                else
                {
                    MessageBox.Show("Xóa thư mục thất bại.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
    }

    private async void MnuMarkAllRead_Click(object sender, RoutedEventArgs e)
    {
        var node = GetNodeFromSender(sender);
        if (node != null)
        {
            TxtStatus.Text = $"Đang đánh dấu đã đọc tất cả thư trong: {node.Name}...";
            if (node.IsOffice365)
            {
                bool success = await Office365Service.MarkAllAsReadAsync(node.Id);
                if (success)
                {
                    var cached = OfflineCacheService.LoadEmails(node.Id);
                    foreach (var email in cached)
                    {
                        email.IsRead = true;
                    }
                    OfflineCacheService.SaveEmails(cached, node.Id);

                    if (_activeFolderId == node.Id)
                    {
                        foreach (var email in _allEmails)
                        {
                            email.IsRead = true;
                        }
                        LstEmails.Items.Refresh();
                    }

                    node.UnreadCount = 0;
                    TxtStatus.Text = $"Đã đánh dấu đã đọc tất cả trong: {node.Name}";
                }
                else
                {
                    MessageBox.Show("Đánh dấu đã đọc thất bại.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            else
            {
                node.UnreadCount = 0;
                if (_activeFolderId == node.Id)
                {
                    foreach (var email in _allEmails)
                    {
                        email.IsRead = true;
                    }
                    LstEmails.Items.Refresh();
                }
                TxtStatus.Text = $"Đã đánh dấu đã đọc tất cả trong: {node.Name}";
            }
        }
    }

    private void MnuChangeFolderColor_Click(object sender, RoutedEventArgs e)
    {
        if (sender is MenuItem mi)
        {
            string color = mi.Tag?.ToString() ?? "Default";
            TxtStatus.Text = $"Đã đổi màu thư mục thành: {color}";
        }
    }

    private void MnuMoveFolder_Click(object sender, RoutedEventArgs e)
    {
        TxtStatus.Text = "Tính năng di chuyển thư mục đang được xử lý.";
    }

    private void MnuEmptyFolder_Click(object sender, RoutedEventArgs e)
    {
        TxtStatus.Text = "Đã dọn dẹp thư mục.";
    }

    private void MnuAddFavorites_Click(object sender, RoutedEventArgs e)
    {
        var node = GetNodeFromSender(sender);
        if (node != null)
        {
            TxtStatus.Text = $"Đã thêm '{node.Name}' vào danh sách Yêu thích.";
        }
    }

    private void MnuMoveUp_Click(object sender, RoutedEventArgs e)
    {
        TxtStatus.Text = "Đã di chuyển thư mục lên trên.";
    }

    private void MnuMoveDown_Click(object sender, RoutedEventArgs e)
    {
        TxtStatus.Text = "Đã di chuyển thư mục xuống dưới.";
    }

    private void MnuClosePst_Click(object sender, RoutedEventArgs e)
    {
        var node = GetNodeFromSender(sender);
        if (node != null)
        {
            if (_activePstFiles.ContainsKey(node.Id))
            {
                var pstFile = _activePstFiles[node.Id];
                pstFile.Dispose();
                _activePstFiles.Remove(node.Id);
                
                SavePstSessions();
                RefreshFolderTreeView();

                // Clear email list if the deleted PST folder was active
                if (_activeFolderId == node.Id || (_activeFolderId != null && _activeFolderId.StartsWith(node.Id)))
                {
                    _allEmails.Clear();
                    UpdateEmailList();
                }

                if (_activePstFiles.Count == 0)
                {
                    ColFolderTree.Width = new GridLength(0);
                    ColFolderTree.MinWidth = 0;
                    if (Splitter1 != null) Splitter1.Visibility = Visibility.Collapsed;
                    PstFolderPanel.Visibility = Visibility.Collapsed;
                }

                TxtStatus.Text = $"Đã đóng tệp dữ liệu: {Path.GetFileName(node.Id)}";
            }
            else
            {
                MessageBox.Show("Vui lòng chọn thư mục gốc (tên tệp PST) để đóng tệp dữ liệu.", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
    }

    private async void InitializeWebView()
    {
        try
        {
            // Check if WebView2 Runtime is available
            string? version = null;
            try
            {
                version = Microsoft.Web.WebView2.Core.CoreWebView2Environment.GetAvailableBrowserVersionString();
            }
            catch { /* Runtime not found */ }

            if (string.IsNullOrEmpty(version))
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
                return;
            }

            // WebView2 Runtime available — initialize with a dedicated user data folder
            var userDataFolder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "MsgViewer", "WebView2");
            var env = await Microsoft.Web.WebView2.Core.CoreWebView2Environment.CreateAsync(null, userDataFolder);
            await EmailWebView.EnsureCoreWebView2Async(env);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"WebView2 init error: {ex.Message}");
        }
    }

    private bool IsWebSessionActive()
    {
        string sessionFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "o365_session.txt");
        return File.Exists(sessionFile);
    }

    private void SaveWebSessionActive(bool active, string? email = null)
    {
        string sessionFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "o365_session.txt");
        try
        {
            if (active)
            {
                File.WriteAllText(sessionFile, email ?? "active");
            }
            else
            {
                if (File.Exists(sessionFile))
                {
                    File.Delete(sessionFile);
                }
            }
        }
        catch { }
    }

    private string GetSavedWebSessionEmail()
    {
        string sessionFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "o365_session.txt");
        if (File.Exists(sessionFile))
        {
            try
            {
                string email = File.ReadAllText(sessionFile).Trim();
                if (email.Contains("@")) return email;
            }
            catch {}
        }
        return "active";
    }

    private async void InitializeOffice365()
    {
        try
        {
            UpdateO365UIState();
            if (IsWebSessionActive())
            {
                await ShowOutlookWebAsync();
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"O365 silent init failed: {ex.Message}");
        }
    }

    private void UpdateO365UIState()
    {
        bool isActive = IsWebSessionActive();
        if (isActive)
        {
            GridO365SignedOut.Visibility = Visibility.Collapsed;
            GridO365SignedIn.Visibility = Visibility.Visible;

            string email = GetSavedWebSessionEmail();
            if (email != "active")
            {
                TxtO365Name.Text = email;
                TxtO365Email.Text = "Đã đăng nhập";
                TxtO365Avatar.Text = email.Substring(0, 1).ToUpper();
            }
            else
            {
                TxtO365Name.Text = "Outlook Web";
                TxtO365Email.Text = "Đã đăng nhập";
                TxtO365Avatar.Text = "O";
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
        await Task.CompletedTask;
    }

    private async void BtnO365SignIn_Click(object sender, RoutedEventArgs e)
    {
        TxtStatus.Text = "Đang mở Outlook Web...";
        SaveWebSessionActive(true);
        UpdateO365UIState();
        await ShowOutlookWebAsync();
        TxtStatus.Text = "Outlook Web đã sẵn sàng.";
    }

    private async void BtnO365SignOut_Click(object sender, RoutedEventArgs e)
    {
        if (MessageBox.Show("Bạn có muốn đăng xuất tài khoản Office 365 không?", "Đăng xuất", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
        {
            StopOwaEmailGrabber();
            SaveWebSessionActive(false);
            UpdateO365UIState();

            // Clear Outlook Web cache/cookies so they are logged out
            try
            {
                if (OutlookWebView.CoreWebView2 != null)
                {
                    await OutlookWebView.CoreWebView2.Profile.ClearBrowsingDataAsync();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error clearing browsing data: {ex.Message}");
            }

            // Hide Outlook Web overlay and restore local layout
            HideOutlookWeb();

            // Clear all email data
            _allEmails.Clear();
            _currentEmail = null;
            _activeFolderId = null;
            UpdateEmailList();

            // Clear email preview WebView
            try
            {
                if (EmailWebView.CoreWebView2 != null)
                {
                    EmailWebView.NavigateToString("<html><body></body></html>");
                }
            }
            catch { }

            TxtStatus.Text = "Đã đăng xuất tài khoản.";
        }
    }

    private async Task ShowOutlookWebAsync()
    {
        try
        {
            // Initialize OutlookWebView if not already done
            if (OutlookWebView.CoreWebView2 == null)
            {
                var userDataFolder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "MsgViewer", "OutlookWebView2");
                var env = await Microsoft.Web.WebView2.Core.CoreWebView2Environment.CreateAsync(null, userDataFolder);
                await OutlookWebView.EnsureCoreWebView2Async(env);
            }

            // Navigate to Outlook Web
            if (OutlookWebView.CoreWebView2 != null)
            {
                OutlookWebView.CoreWebView2.Navigate("https://outlook.office.com/mail/");
            }

            // Show and select Outlook Web tab
            if (RadOutlookWeb != null)
            {
                RadOutlookWeb.Visibility = Visibility.Visible;
                RadOutlookWeb.IsChecked = true;
            }

            StartOwaEmailGrabber();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Không thể tải Outlook Web: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void HideOutlookWeb()
    {
        StopOwaEmailGrabber();

        // Select and hide Outlook Web tab
        if (RadLocalViewer != null)
        {
            RadLocalViewer.IsChecked = true;
        }
        if (RadOutlookWeb != null)
        {
            RadOutlookWeb.Visibility = Visibility.Collapsed;
        }

        // Navigate away to release resources
        try
        {
            if (OutlookWebView.CoreWebView2 != null)
            {
                OutlookWebView.CoreWebView2.Navigate("about:blank");
            }
        }
        catch { }

        // Reset sidebar (keep open if active PSTs are present)
        if (_activePstFiles.Count > 0)
        {
            ColFolderTree.Width = new GridLength(240);
            ColFolderTree.MinWidth = 150;
            if (Splitter1 != null) Splitter1.Visibility = Visibility.Visible;
            PstFolderPanel.Visibility = Visibility.Visible;
            RefreshFolderTreeView();
        }
        else
        {
            TreePstFolders.ItemsSource = null;
            ColFolderTree.Width = new GridLength(0);
            ColFolderTree.MinWidth = 0;
            if (Splitter1 != null) Splitter1.Visibility = Visibility.Collapsed;
            PstFolderPanel.Visibility = Visibility.Collapsed;
        }
    }

    private void StartOwaEmailGrabber()
    {
        if (_owaEmailGrabberTimer == null)
        {
            _owaEmailGrabberTimer = new System.Windows.Threading.DispatcherTimer();
            _owaEmailGrabberTimer.Interval = TimeSpan.FromSeconds(3);
            _owaEmailGrabberTimer.Tick += OwaEmailGrabberTimer_Tick;
        }
        _owaEmailGrabberTimer.Start();
    }

    private void StopOwaEmailGrabber()
    {
        _owaEmailGrabberTimer?.Stop();
    }

    private async void OwaEmailGrabberTimer_Tick(object? sender, EventArgs e)
    {
        await GrabOwaEmailAsync();
    }

    private async Task GrabOwaEmailAsync()
    {
        if (OutlookWebView.CoreWebView2 == null || OutlookWebOverlay.Visibility != Visibility.Visible) return;

        try
        {
            string js = @"
(function() {
    try {
        if (window.sessionContext && window.sessionContext.UserEmailAddress) {
            return window.sessionContext.UserEmailAddress;
        }
        if (window.Owa && window.Owa.Configuration && window.Owa.Configuration.UserEmailAddress) {
            return window.Owa.Configuration.UserEmailAddress;
        }
        if (window.g_PersonaLayouts && window.g_PersonaLayouts.userEmail) {
            return window.g_PersonaLayouts.userEmail;
        }
        const elems = document.querySelectorAll('[title], [aria-label]');
        const emailRegex = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/;
        for (let i = 0; i < elems.length; i++) {
            const el = elems[i];
            const title = el.getAttribute('title') || '';
            const ariaLabel = el.getAttribute('aria-label') || '';
            
            let match = title.match(emailRegex);
            if (match) return match[0];
            
            match = ariaLabel.match(emailRegex);
            if (match) return match[0];
        }
    } catch(e) {}
    return null;
})()";
            string resultJson = await OutlookWebView.CoreWebView2.ExecuteScriptAsync(js);
            if (!string.IsNullOrEmpty(resultJson) && resultJson != "null")
            {
                string email = resultJson.Trim('\"');
                if (email.Contains("@") && email.Length > 5)
                {
                    SaveWebSessionActive(true, email);
                    UpdateO365UIState();
                }
            }
        }
        catch {}
    }

    private void RadOutlookWeb_Checked(object sender, RoutedEventArgs e)
    {
        if (OutlookWebOverlay != null && MainLayoutGrid != null)
        {
            OutlookWebOverlay.Visibility = Visibility.Visible;
            MainLayoutGrid.Visibility = Visibility.Collapsed;
        }
    }

    private void RadLocalViewer_Checked(object sender, RoutedEventArgs e)
    {
        if (OutlookWebOverlay != null && MainLayoutGrid != null)
        {
            OutlookWebOverlay.Visibility = Visibility.Collapsed;
            MainLayoutGrid.Visibility = Visibility.Visible;
        }
    }

    private void SwitchToLocalViewer()
    {
        if (RadLocalViewer != null && RadLocalViewer.IsChecked == false)
        {
            RadLocalViewer.IsChecked = true;
        }
    }

    private void BtnO365Compose_Click(object sender, RoutedEventArgs e)
    {
        var composeWin = new ComposeWindow { Owner = this };
        composeWin.ShowDialog();
    }

    private void BtnRibbonReport_Click(object sender, RoutedEventArgs e)
    {
        if (_currentEmail != null && _currentEmail.FilePath.StartsWith("o365://"))
        {
            MenuJunk_Click(sender, e);
        }
        else
        {
            MessageBox.Show("Vui lòng chọn một email Office 365 để báo cáo thư rác.", "Báo cáo thư rác", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }

    private void BtnRibbonSweep_Click(object sender, RoutedEventArgs e)
    {
        MessageBox.Show("Tính năng dọn dẹp thư (Sweep) giúp tự động di chuyển hoặc xóa các email cũ từ người gửi này. Tính năng đang được thiết lập.", "Sweep", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private void BtnRibbonMoveTo_Click(object sender, RoutedEventArgs e)
    {
        MessageBox.Show("Tính năng di chuyển thư (Move to) giúp phân loại email vào các thư mục lưu trữ khác nhau.", "Move to", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private void BtnRibbonShareTeams_Click(object sender, RoutedEventArgs e)
    {
        MessageBox.Show("Đang mở liên kết chia sẻ email qua Microsoft Teams...", "Share to Teams", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private void BtnReply_Click(object sender, RoutedEventArgs e) => ActivateInlineEditor("reply");

    private void BtnReplyAll_Click(object sender, RoutedEventArgs e) => ActivateInlineEditor("replyall");

    private void BtnForward_Click(object sender, RoutedEventArgs e) => ActivateInlineEditor("forward");

    private void BtnQuickReply_Click(object sender, RoutedEventArgs e) => ActivateInlineEditor("reply");

    private void BtnQuickReplyAll_Click(object sender, RoutedEventArgs e) => ActivateInlineEditor("replyall");

    private void BtnQuickForward_Click(object sender, RoutedEventArgs e) => ActivateInlineEditor("forward");

    private void ActivateInlineEditor(string actionType)
    {
        if (_currentEmail == null) return;
        
        _inlineActionType = actionType;
        _inlineAttachments.Clear();
        RefreshInlineAttachmentsList();

        RtfInlineBody.Document.Blocks.Clear();
        TxtInlineDraftStatus.Text = "";

        BdrInlineEditor.Visibility = Visibility.Visible;

        if (actionType == "reply")
        {
            TxtInlineActionType.Text = "↩️ Replying to sender...";
            TxtInlineTo.Text = _currentEmail.FromEmail + "; ";
            TxtInlineCc.Text = "";
        }
        else if (actionType == "replyall")
        {
            TxtInlineActionType.Text = "🔂 Replying to all...";
            var recipients = new List<string>();
            if (!string.IsNullOrEmpty(_currentEmail.FromEmail)) recipients.Add(_currentEmail.FromEmail);
            if (!string.IsNullOrEmpty(_currentEmail.To)) recipients.Add(_currentEmail.To);
            
            TxtInlineTo.Text = string.Join("; ", recipients.Distinct()) + "; ";
            TxtInlineCc.Text = _currentEmail.Cc ?? "";
        }
        else if (actionType == "forward")
        {
            TxtInlineActionType.Text = "➡️ Forwarding email...";
            TxtInlineTo.Text = "";
            TxtInlineCc.Text = "";

            // Copy attachments from original email
            foreach (var att in _currentEmail.Attachments)
            {
                try
                {
                    string tempPath = Path.Combine(Path.GetTempPath(), att.FileName);
                    File.WriteAllBytes(tempPath, att.Data);
                    _inlineAttachments.Add(tempPath);
                }
                catch {}
            }
            RefreshInlineAttachmentsList();
        }

        StartInlineDraftTimer();
    }

    private void StartInlineDraftTimer()
    {
        if (_inlineDraftTimer == null)
        {
            _inlineDraftTimer = new System.Windows.Threading.DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(15)
            };
            _inlineDraftTimer.Tick += InlineDraftTimer_Tick;
        }
        _inlineDraftId = null;
        _inlineDraftTimer.Start();
    }

    private void StopInlineDraftTimer()
    {
        if (_inlineDraftTimer != null)
        {
            _inlineDraftTimer.Stop();
        }
        _inlineDraftId = null;
    }

    private async void InlineDraftTimer_Tick(object? sender, EventArgs e)
    {
        if (BdrInlineEditor.Visibility != Visibility.Visible || _currentEmail == null) return;

        try
        {
            TxtInlineDraftStatus.Text = "Đang lưu nháp ngầm...";

            string toStr = TxtInlineTo.Text;
            string ccStr = TxtInlineCc.Text;
            string subject = _inlineActionType == "reply" || _inlineActionType == "replyall"
                ? "Re: " + _currentEmail.Subject
                : "Fw: " + _currentEmail.Subject;

            string htmlBody = ConvertFlowDocumentToHtml(RtfInlineBody.Document);

            // Quote original message in draft
            string quotedHtml = htmlBody + "<br/><br/><hr/>" + (_currentEmail.BodyHtml ?? "");

            var draftId = await Office365Service.CreateOrUpdateDraftAsync(
                _inlineDraftId,
                toStr,
                ccStr,
                subject,
                quotedHtml
            );

            if (!string.IsNullOrEmpty(draftId))
            {
                _inlineDraftId = draftId;
                TxtInlineDraftStatus.Text = $"Đã tự động lưu nháp lúc {DateTime.Now:HH:mm:ss}";
            }
            else
            {
                TxtInlineDraftStatus.Text = "Không thể lưu nháp.";
            }
        }
        catch (Exception ex)
        {
            TxtInlineDraftStatus.Text = "Lỗi lưu nháp.";
            System.Diagnostics.Debug.WriteLine($"Inline draft save error: {ex.Message}");
        }
    }

    private async void BtnInlineSend_Click(object sender, RoutedEventArgs e)
    {
        if (_currentEmail == null) return;

        try
        {
            TxtStatus.Text = "Đang gửi email...";
            StopInlineDraftTimer();

            string toStr = TxtInlineTo.Text;
            string ccStr = TxtInlineCc.Text;
            string subject = _inlineActionType == "reply" || _inlineActionType == "replyall"
                ? "Re: " + _currentEmail.Subject
                : "Fw: " + _currentEmail.Subject;

            string htmlBody = ConvertFlowDocumentToHtml(RtfInlineBody.Document);

            // Quote original message
            string quotedHtml = htmlBody + "<br/><br/><hr/>" + (_currentEmail.BodyHtml ?? "");

            bool success = await Office365Service.SendEmailAsync(
                toStr,
                ccStr,
                subject,
                quotedHtml,
                _inlineAttachments
            );

            if (success)
            {
                if (!string.IsNullOrEmpty(_inlineDraftId))
                {
                    try
                    {
                        await Office365Service.DeleteEmailAsync(_inlineDraftId);
                    }
                    catch {}
                }

                MessageBox.Show("Đã gửi email phản hồi thành công!", "Thành công", MessageBoxButton.OK, MessageBoxImage.Information);
                TxtStatus.Text = "Đã gửi thư.";
                
                BdrInlineEditor.Visibility = Visibility.Collapsed;
            }
            else
            {
                MessageBox.Show("Gửi email thất bại. Vui lòng kiểm tra lại kết nối mạng.", "Thất bại", MessageBoxButton.OK, MessageBoxImage.Warning);
                TxtStatus.Text = "Gửi thư thất bại.";
                StartInlineDraftTimer();
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Lỗi gửi thư: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            TxtStatus.Text = "Lỗi gửi thư.";
            StartInlineDraftTimer();
        }
    }

    private async void BtnInlineDiscard_Click(object sender, RoutedEventArgs e)
    {
        if (MessageBox.Show("Bạn có chắc muốn hủy bản soạn thảo nháp này không?", "Hủy soạn thảo", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
        {
            StopInlineDraftTimer();

            if (!string.IsNullOrEmpty(_inlineDraftId))
            {
                try
                {
                    TxtStatus.Text = "Đang xóa thư nháp trên server...";
                    await Office365Service.DeleteEmailAsync(_inlineDraftId);
                }
                catch {}
            }

            TxtStatus.Text = "Đã hủy soạn thảo.";
            BdrInlineEditor.Visibility = Visibility.Collapsed;
        }
    }

    private void BtnInlineBold_Click(object sender, RoutedEventArgs e)
    {
        if (RtfInlineBody.Selection != null)
        {
            var textWeight = RtfInlineBody.Selection.GetPropertyValue(TextElement.FontWeightProperty);
            if (textWeight is FontWeight weight && weight == FontWeights.Bold)
            {
                RtfInlineBody.Selection.ApplyPropertyValue(TextElement.FontWeightProperty, FontWeights.Normal);
            }
            else
            {
                RtfInlineBody.Selection.ApplyPropertyValue(TextElement.FontWeightProperty, FontWeights.Bold);
            }
        }
    }

    private void BtnInlineItalic_Click(object sender, RoutedEventArgs e)
    {
        if (RtfInlineBody.Selection != null)
        {
            var textStyle = RtfInlineBody.Selection.GetPropertyValue(TextElement.FontStyleProperty);
            if (textStyle is FontStyle style && style == FontStyles.Italic)
            {
                RtfInlineBody.Selection.ApplyPropertyValue(TextElement.FontStyleProperty, FontStyles.Normal);
            }
            else
            {
                RtfInlineBody.Selection.ApplyPropertyValue(TextElement.FontStyleProperty, FontStyles.Italic);
            }
        }
    }

    private void BtnInlineUnderline_Click(object sender, RoutedEventArgs e)
    {
        if (RtfInlineBody.Selection != null)
        {
            var textDecoration = RtfInlineBody.Selection.GetPropertyValue(Inline.TextDecorationsProperty);
            if (textDecoration != DependencyProperty.UnsetValue && textDecoration != null)
            {
                RtfInlineBody.Selection.ApplyPropertyValue(Inline.TextDecorationsProperty, null);
            }
            else
            {
                RtfInlineBody.Selection.ApplyPropertyValue(Inline.TextDecorationsProperty, TextDecorations.Underline);
            }
        }
    }

    private void BtnInlineBlueText_Click(object sender, RoutedEventArgs e)
    {
        if (RtfInlineBody.Selection != null)
        {
            RtfInlineBody.Selection.ApplyPropertyValue(TextElement.ForegroundProperty, new SolidColorBrush((Color)ColorConverter.ConvertFromString("#0078d4")));
        }
    }

    private void BtnInlineAttach_Click(object sender, RoutedEventArgs e)
    {
        var openFileDialog = new Microsoft.Win32.OpenFileDialog
        {
            Multiselect = true,
            Title = "Chọn tệp tin đính kèm"
        };
        if (openFileDialog.ShowDialog() == true)
        {
            foreach (var filename in openFileDialog.FileNames)
            {
                if (!_inlineAttachments.Contains(filename))
                {
                    _inlineAttachments.Add(filename);
                }
            }
            RefreshInlineAttachmentsList();
        }
    }

    private void RtfInlineBody_PreviewDragOver(object sender, DragEventArgs e)
    {
        if (e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            e.Effects = DragDropEffects.Copy;
            e.Handled = true;
        }
    }

    private void RtfInlineBody_Drop(object sender, DragEventArgs e)
    {
        if (e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (files != null)
            {
                foreach (string file in files)
                {
                    if (File.Exists(file))
                    {
                        if (!_inlineAttachments.Contains(file))
                        {
                            _inlineAttachments.Add(file);
                        }
                    }
                }
                RefreshInlineAttachmentsList();
            }
        }
    }

    private void RefreshInlineAttachmentsList()
    {
        LstInlineAttachments.ItemsSource = null;
        LstInlineAttachments.ItemsSource = _inlineAttachments.Select(f => new { FileName = Path.GetFileName(f), FilePath = f }).ToList();
    }

    private void BtnInlineRemoveAttachment_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button btn && btn.Tag is string path)
        {
            _inlineAttachments.Remove(path);
            RefreshInlineAttachmentsList();
        }
    }

    private string ConvertFlowDocumentToHtml(System.Windows.Documents.FlowDocument doc)
    {
        var sb = new System.Text.StringBuilder();
        sb.Append("<html><body style='font-family:Segoe UI,Arial; font-size:10.5pt;'>");

        foreach (var block in doc.Blocks)
        {
            if (block is System.Windows.Documents.Paragraph p)
            {
                sb.Append("<p>");
                foreach (var inline in p.Inlines)
                {
                    sb.Append(ConvertInlineToHtml(inline));
                }
                sb.Append("</p>");
            }
            else if (block is System.Windows.Documents.List list)
            {
                sb.Append("<ul>");
                foreach (var item in list.ListItems)
                {
                    sb.Append("<li>");
                    foreach (var listBlock in item.Blocks)
                    {
                        if (listBlock is System.Windows.Documents.Paragraph lp)
                        {
                            foreach (var inline in lp.Inlines)
                            {
                                sb.Append(ConvertInlineToHtml(inline));
                            }
                        }
                    }
                    sb.Append("</li>");
                }
                sb.Append("</ul>");
            }
        }

        sb.Append("</body></html>");
        return sb.ToString();
    }

    private string ConvertInlineToHtml(System.Windows.Documents.Inline inline)
    {
        if (inline is System.Windows.Documents.Run run)
        {
            string text = System.Net.WebUtility.HtmlEncode(run.Text);
            bool bold = run.FontWeight == FontWeights.Bold;
            bool italic = run.FontStyle == FontStyles.Italic;
            bool underline = run.TextDecorations != null && run.TextDecorations.Count > 0;

            if (bold) text = $"<strong>{text}</strong>";
            if (italic) text = $"<em>{text}</em>";
            if (underline) text = $"<u>{text}</u>";

            var fg = run.Foreground as System.Windows.Media.SolidColorBrush;
            if (fg != null && fg.Color != System.Windows.Media.Colors.Black)
            {
                string colorHex = fg.Color.ToString().Substring(3); // Remove Alpha channel
                text = $"<span style='color:#{colorHex};'>{text}</span>";
            }

            return text;
        }
        else if (inline is System.Windows.Documents.LineBreak)
        {
            return "<br/>";
        }
        return "";
    }

    private async void BtnO365Archive_Click(object sender, RoutedEventArgs e)
    {
        if (_currentEmail == null || !_currentEmail.FilePath.StartsWith("o365://")) return;

        try
        {
            TxtStatus.Text = "Đang lưu trữ thư...";
            string messageId = _currentEmail.FilePath.Substring("o365://".Length);
            await Office365Service.ArchiveEmailAsync(messageId);

            var removedEmail = _currentEmail;
            _allEmails.Remove(removedEmail);
            UpdateEmailList();

            GridPlaceholder.Visibility = Visibility.Visible;
            GridDetail.Visibility = Visibility.Collapsed;
            _currentEmail = null;

            TxtStatus.Text = $"Đã di chuyển thư \"{removedEmail.Subject}\" sang mục lưu trữ.";
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Lỗi lưu trữ thư: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            TxtStatus.Text = "Lỗi khi lưu trữ thư.";
        }
    }

    private async void BtnO365Delete_Click(object sender, RoutedEventArgs e)
    {
        if (_currentEmail == null || !_currentEmail.FilePath.StartsWith("o365://")) return;

        if (MessageBox.Show("Bạn có chắc chắn muốn xóa email này không?", "Xác nhận xóa", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
        {
            try
            {
                TxtStatus.Text = "Đang xóa thư...";
                string messageId = _currentEmail.FilePath.Substring("o365://".Length);
                await Office365Service.DeleteEmailAsync(messageId);

                var removedEmail = _currentEmail;
                _allEmails.Remove(removedEmail);
                UpdateEmailList();

                GridPlaceholder.Visibility = Visibility.Visible;
                GridDetail.Visibility = Visibility.Collapsed;
                _currentEmail = null;

                TxtStatus.Text = $"Đã di chuyển thư \"{removedEmail.Subject}\" sang mục đã xóa.";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi xóa thư: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                TxtStatus.Text = "Lỗi khi xóa thư.";
            }
        }
    }

    private async void BtnQuickToggleRead_Click(object sender, RoutedEventArgs e)
    {
        var email = (sender as Button)?.DataContext as EmailMessage;
        if (email == null) return;

        try
        {
            if (email.FilePath.StartsWith("o365://"))
            {
                string messageId = email.FilePath.Substring("o365://".Length);
                bool nextState = !email.IsRead;
                await Office365Service.MarkAsReadAsync(messageId, nextState);
                email.IsRead = nextState;
                LstEmails.Items.Refresh();
                TxtStatus.Text = nextState ? "Đã đánh dấu là Đã đọc." : "Đã đánh dấu là Chưa đọc.";
            }
            else
            {
                email.IsRead = !email.IsRead;
                LstEmails.Items.Refresh();
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Lỗi thay đổi trạng thái đọc: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async void BtnQuickDelete_Click(object sender, RoutedEventArgs e)
    {
        var email = (sender as Button)?.DataContext as EmailMessage;
        if (email == null) return;

        if (MessageBox.Show("Bạn có chắc chắn muốn xóa email này không?", "Xác nhận xóa", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
        {
            try
            {
                if (email.FilePath.StartsWith("o365://"))
                {
                    TxtStatus.Text = "Đang xóa thư...";
                    string messageId = email.FilePath.Substring("o365://".Length);
                    await Office365Service.DeleteEmailAsync(messageId);
                }

                _allEmails.Remove(email);
                UpdateEmailList();

                if (_currentEmail == email)
                {
                    GridPlaceholder.Visibility = Visibility.Visible;
                    GridDetail.Visibility = Visibility.Collapsed;
                    _currentEmail = null;
                }

                TxtStatus.Text = $"Đã xóa email: {email.Subject}";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi xóa thư: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }

    private void BtnRibbonDelete_Click(object sender, RoutedEventArgs e)
    {
        BtnO365Delete_Click(sender, e);
    }

    private void BtnRibbonArchive_Click(object sender, RoutedEventArgs e)
    {
        BtnO365Archive_Click(sender, e);
    }

    private async void BtnRibbonMarkRead_Click(object sender, RoutedEventArgs e)
    {
        if (_currentEmail == null || !_currentEmail.FilePath.StartsWith("o365://")) return;

        try
        {
            string messageId = _currentEmail.FilePath.Substring("o365://".Length);
            bool nextState = !_currentEmail.IsRead;
            TxtStatus.Text = nextState ? "Đang đánh dấu đã đọc..." : "Đang đánh dấu chưa đọc...";
            await Office365Service.MarkAsReadAsync(messageId, nextState);
            _currentEmail.IsRead = nextState;
            LstEmails.Items.Refresh();
            TxtStatus.Text = nextState ? "Đã đánh dấu là Đã đọc." : "Đã đánh dấu là Chưa đọc.";
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Lỗi thay đổi trạng thái đọc: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void BtnFileMenu_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button btn)
        {
            btn.ContextMenu.IsOpen = true;
        }
    }

    private void MnuFileOpenMsgEml_Click(object sender, RoutedEventArgs e)
    {
        BtnOpenFile_Click(sender, e);
    }

    private void MnuFileOpenPstOst_Click(object sender, RoutedEventArgs e)
    {
        BtnOpenPst_Click(sender, e);
    }

    private async void MnuFileImportPst_Click(object sender, RoutedEventArgs e)
    {
        if (!Office365Service.IsSignedIn)
        {
            MessageBox.Show("Vui lòng đăng nhập tài khoản Office 365 để thực hiện tính năng này.", "Yêu cầu đăng nhập", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var dialog = new Microsoft.Win32.OpenFileDialog
        {
            Filter = "Outlook Data Files (*.pst)|*.pst|Offline Storage Table (*.ost)|*.ost",
            Title = "Chọn tệp PST để nhập vào Office 365"
        };

        if (dialog.ShowDialog() == true)
        {
            string filePath = dialog.FileName;
            
            if (MessageBox.Show($"Bạn có chắc chắn muốn nhập dữ liệu từ file \"{Path.GetFileName(filePath)}\" vào tài khoản Office 365 của bạn không?\nQuá trình này sẽ tải tất cả thư mục và email lên đám mây.", "Nhập dữ liệu PST", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                try
                {
                    TxtStatus.Text = "Đang bắt đầu nhập PST...";
                    var progressWin = new Window
                    {
                        Title = "Đang nhập dữ liệu PST vào Office 365",
                        Width = 400,
                        Height = 150,
                        WindowStartupLocation = WindowStartupLocation.CenterOwner,
                        Owner = this,
                        ResizeMode = ResizeMode.NoResize
                    };
                    
                    var sp = new StackPanel { Margin = new Thickness(20), VerticalAlignment = VerticalAlignment.Center };
                    var lblStatus = new TextBlock { Text = "Đang quét tệp dữ liệu...", FontWeight = FontWeights.SemiBold, Margin = new Thickness(0,0,0,10) };
                    var pb = new ProgressBar { Height = 18, Minimum = 0, Maximum = 100, IsIndeterminate = true };
                    sp.Children.Add(lblStatus);
                    sp.Children.Add(pb);
                    progressWin.Content = sp;
                    progressWin.Show();

                    await System.Threading.Tasks.Task.Run(async () =>
                    {
                        try
                        {
                            var xstFile = new XstFile(filePath);
                            var stats = new PstImportStats();

                            // Recursively import folders
                            await ImportFolderRecursivelyAsync(xstFile.RootFolder, null, (status) =>
                            {
                                Dispatcher.Invoke(() =>
                                {
                                    lblStatus.Text = status;
                                });
                            }, stats);

                            Dispatcher.Invoke(() =>
                            {
                                progressWin.Close();
                                MessageBox.Show($"Nhập PST thành công!\nĐã nhập {stats.FoldersCount} thư mục và {stats.EmailsCount} email vào Office 365.", "Thành công", MessageBoxButton.OK, MessageBoxImage.Information);
                                TxtStatus.Text = "Nhập dữ liệu PST hoàn tất.";
                            });
                        }
                        catch (Exception ex)
                        {
                            Dispatcher.Invoke(() =>
                            {
                                progressWin.Close();
                                MessageBox.Show($"Lỗi trong quá trình nhập dữ liệu: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                                TxtStatus.Text = "Lỗi nhập PST.";
                            });
                        }
                    });
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Không thể khởi động quá trình nhập PST: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
    }

    private class PstImportStats
    {
        public int FoldersCount { get; set; }
        public int EmailsCount { get; set; }
    }

    private async System.Threading.Tasks.Task ImportFolderRecursivelyAsync(XstFolder folder, string? parentFolderId, Action<string> progressCallback, PstImportStats stats)
    {
        if (folder == null) return;

        string folderName = folder.DisplayName ?? "";
        if (string.IsNullOrWhiteSpace(folderName)) return;

        progressCallback($"Đang tạo thư mục: {folderName}...");

        string? createdFolderId = null;
        try
        {
            createdFolderId = await Office365Service.CreateMailFolderAsync(folderName, parentFolderId);
            stats.FoldersCount++;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error creating folder {folderName}: {ex.Message}");
        }

        // Upload messages in this folder
        var messages = folder.Messages ?? folder.GetMessages();
        if (messages != null)
        {
            var msgList = messages.ToList();
            for (int i = 0; i < msgList.Count; i++)
            {
                var msg = msgList[i];
                stats.EmailsCount++;
                progressCallback($"[{folderName}] Đang tải thư {i+1}/{msgList.Count}...");

                try
                {
                    var tempEmail = new EmailMessage { RawXstMessage = msg };
                    PstParser.LoadMessageDetails(tempEmail);
                    
                    string subject = string.IsNullOrWhiteSpace(msg.Subject) ? "(Không có tiêu đề)" : msg.Subject;
                    string htmlBody = tempEmail.BodyHtml ?? tempEmail.BodyText ?? "";
                    string toStr = msg.To ?? "";
                    DateTime? msgDate = msg.Date ?? msg.ReceivedTime ?? msg.SubmittedTime;

                    await Office365Service.UploadImportedEmailAsync(
                        createdFolderId,
                        toStr,
                        subject,
                        htmlBody,
                        msgDate
                    );
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error uploading message: {ex.Message}");
                }
            }
        }

        // Process subfolders
        if (folder.Folders != null)
        {
            foreach (var subFolder in folder.Folders)
            {
                await ImportFolderRecursivelyAsync(subFolder, createdFolderId, progressCallback, stats);
            }
        }
    }

    private void BtnOpenFile_Click(object sender, RoutedEventArgs e)
    {
        SwitchToLocalViewer();
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
        SwitchToLocalViewer();
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


    private void FilterTab_Checked(object sender, RoutedEventArgs e)
    {
        if (RadFilterUnread == null) return;
        _showUnreadOnlyTab = RadFilterUnread.IsChecked == true;
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
            ParseOutlookSearchQuery(keyword, out string? filterSubject, out string? filterFrom, out string? filterTo, out bool? filterHasAttachment, out bool? filterIsUnread, out string? generalKeyword);

            SearchKeyword = generalKeyword ?? keyword;

            if (filterFrom != null)
            {
                emailsToFilter = emailsToFilter.Where(e => IsMatch(e.FromName, filterFrom) || IsMatch(e.FromEmail, filterFrom));
            }
            if (filterSubject != null)
            {
                emailsToFilter = emailsToFilter.Where(e => IsMatch(e.Subject, filterSubject));
            }
            if (filterTo != null)
            {
                emailsToFilter = emailsToFilter.Where(e => IsMatch(e.To, filterTo));
            }
            if (filterHasAttachment.HasValue)
            {
                emailsToFilter = emailsToFilter.Where(e => e.HasAttachments == filterHasAttachment.Value);
            }
            if (filterIsUnread.HasValue)
            {
                emailsToFilter = emailsToFilter.Where(e => !e.IsRead == filterIsUnread.Value);
            }

            if (generalKeyword != null)
            {
                int scopeIndex = CboSearchScope?.SelectedIndex ?? 0;
                emailsToFilter = emailsToFilter.Where(e => {
                    switch (scopeIndex)
                    {
                        case 1: // Subject only
                            return IsMatch(e.Subject, generalKeyword);
                        case 2: // Sender only
                            return IsMatch(e.FromName, generalKeyword) || IsMatch(e.FromEmail, generalKeyword);
                        case 3: // Recipient only
                            return IsMatch(e.To, generalKeyword);
                        case 0: // All fields
                        default:
                            return IsMatch(e.Subject, generalKeyword) ||
                                   IsMatch(e.FromName, generalKeyword) ||
                                   IsMatch(e.FromEmail, generalKeyword) ||
                                   IsMatch(e.To, generalKeyword) ||
                                   IsMatch(e.BodyText, generalKeyword);
                    }
                });
            }
        }
        else
        {
            SearchKeyword = "";
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

        // Apply Unread Filtering
        if (ChkUnreadOnly != null && ChkUnreadOnly.IsChecked == true)
        {
            emailsToFilter = emailsToFilter.Where(e => !e.IsRead);
        }

        // Apply Sorting
        if (_currentSortTag == "DateAsc")
        {
            emailsToFilter = emailsToFilter.OrderBy(e => e.Date ?? DateTime.MinValue);
        }
        else if (_currentSortTag == "FromAsc")
        {
            emailsToFilter = emailsToFilter.OrderBy(e => e.FromName ?? e.FromEmail ?? "");
        }
        else if (_currentSortTag == "FromDesc")
        {
            emailsToFilter = emailsToFilter.OrderByDescending(e => e.FromName ?? e.FromEmail ?? "");
        }
        else if (_currentSortTag == "SubjectAsc")
        {
            emailsToFilter = emailsToFilter.OrderBy(e => e.Subject ?? "");
        }
        else if (_currentSortTag == "SubjectDesc")
        {
            emailsToFilter = emailsToFilter.OrderByDescending(e => e.Subject ?? "");
        }
        else // DateDesc (default)
        {
            emailsToFilter = emailsToFilter.OrderByDescending(e => e.Date ?? DateTime.MinValue);
        }

        // Apply Unread Tab filter
        if (_showUnreadOnlyTab)
        {
            emailsToFilter = emailsToFilter.Where(e => !e.IsRead);
        }

        var filteredList = emailsToFilter.ToList();
        LstEmails.ItemsSource = filteredList;
        BdrEmptyList.Visibility = filteredList.Any() ? Visibility.Collapsed : Visibility.Visible;

        var view = System.Windows.Data.CollectionViewSource.GetDefaultView(LstEmails.ItemsSource);
        if (view != null)
        {
            view.GroupDescriptions.Clear();
            view.GroupDescriptions.Add(new System.Windows.Data.PropertyGroupDescription("DateGroup"));
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
            if (_activeFolderId == "drafts" && selectedEmail.FilePath.StartsWith("o365://"))
            {
                LstEmails.SelectionChanged -= LstEmails_SelectionChanged;
                LstEmails.SelectedItem = null;
                LstEmails.SelectionChanged += LstEmails_SelectionChanged;

                var composeWin = new ComposeWindow(selectedEmail, "draft") { Owner = this };
                if (composeWin.ShowDialog() == true)
                {
                    RefreshCurrentFolder();
                }
            }
            else
            {
                LoadEmailDetails(selectedEmail);
                if (!selectedEmail.IsRead)
                {
                    selectedEmail.IsRead = true;
                    LstEmails.Items.Refresh();

                    // If O365, call remote API in background to sync
                    if (selectedEmail.FilePath.StartsWith("o365://"))
                    {
                        string messageId = selectedEmail.FilePath.Substring("o365://".Length);
                        _ = Office365Service.MarkAsReadAsync(messageId, true);
                    }

                    // Decrement TreeView selected folder node unread count
                    if (TreePstFolders.SelectedItem is PstFolderNode node)
                    {
                        if (node.UnreadCount > 0)
                        {
                            node.UnreadCount--;
                        }
                    }
                }
            }
        }
        else
        {
            GridPlaceholder.Visibility = Visibility.Visible;
            GridDetail.Visibility = Visibility.Collapsed;
        }
    }

    private void RefreshCurrentFolder()
    {
        if (TreePstFolders.SelectedItem is PstFolderNode node)
        {
            TreePstFolders_SelectedItemChanged(TreePstFolders, new RoutedPropertyChangedEventArgs<object>(node, node));
        }
    }


    private async void LoadEmailDetails(EmailMessage email)
    {
        _currentEmail = email;
        GridPlaceholder.Visibility = Visibility.Collapsed;
        GridDetail.Visibility = Visibility.Visible;

        StopInlineDraftTimer();
        if (BdrInlineEditor != null) BdrInlineEditor.Visibility = Visibility.Collapsed;

        bool isO365 = email.FilePath.StartsWith("o365://");
        if (BtnO365Archive != null) BtnO365Archive.Visibility = isO365 ? Visibility.Visible : Visibility.Collapsed;
        if (BtnO365Delete != null) BtnO365Delete.Visibility = isO365 ? Visibility.Visible : Visibility.Collapsed;

        bool canCompose = Office365Service.IsSignedIn;

        if (BtnDetailReply != null) BtnDetailReply.IsEnabled = canCompose;
        if (BtnDetailReplyAll != null) BtnDetailReplyAll.IsEnabled = canCompose;
        if (BtnDetailForward != null) BtnDetailForward.IsEnabled = canCompose;

        // Nạp lazy body & attachments từ XstMessage nếu là thư trong PST
        if (email.RawXstMessage != null)
        {
            PstParser.LoadMessageDetails(email);
        }
        else if (email.FilePath.StartsWith("o365://"))
        {
            TxtStatus.Text = "Đang tải chi tiết thư từ Office 365...";
            await Office365Service.LoadEmailFullDetailsAsync(email);

            if (!email.IsRead)
            {
                try
                {
                    string messageId = email.FilePath.Substring("o365://".Length);
                    await Office365Service.MarkAsReadAsync(messageId, true);
                    email.IsRead = true;
                    LstEmails.Items.Refresh();
                }
                catch (Exception readEx)
                {
                    System.Diagnostics.Debug.WriteLine($"Lỗi đánh dấu đã đọc: {readEx.Message}");
                }
            }
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
        SwitchToLocalViewer();
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
        SwitchToLocalViewer();
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

            if (_activePstFiles.ContainsKey(filePath))
            {
                TxtStatus.Text = $"Tệp dữ liệu đã được mở: {Path.GetFileName(filePath)}";
                return;
            }

            var pstFile = new XstFile(filePath);
            _activePstFiles[filePath] = pstFile;

            // Cấu hình giao diện chế độ 3 cột
            ColFolderTree.Width = new GridLength(240);
            ColFolderTree.MinWidth = 150;
            if (Splitter1 != null) Splitter1.Visibility = Visibility.Visible;
            PstFolderPanel.Visibility = Visibility.Visible;

            SavePstSessions();
            RefreshFolderTreeView();

            TxtStatus.Text = $"Đã mở thành công tệp dữ liệu Outlook: {Path.GetFileName(filePath)}";
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Lỗi khi mở tệp PST: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            TxtStatus.Text = "Lỗi khi nạp tệp PST.";
        }
    }

    private PstFolderNode CreateFolderNode(XstFolder folder, string pstFilePath)
    {
        var node = new PstFolderNode
        {
            Name = folder.DisplayName ?? "(Thư mục không tên)",
            Folder = folder,
            PstFilePath = pstFilePath,
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
                node.SubFolders.Add(CreateFolderNode(sub, pstFilePath));
            }
        }
        return node;
    }

    private async void TreePstFolders_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
    {
        if (e.NewValue is PstFolderNode node)
        {
            _activeFolderId = node.Id;
            try
            {
                TxtStatus.Text = $"Đang đọc danh sách thư từ thư mục: {node.Name}...";
                _allEmails.Clear();

                if (node.IsOffice365)
                {
                    if (node.Id == "root_account")
                    {
                        TxtStatus.Text = $"Tài khoản: {node.Name}";
                        return;
                    }

                    // 1. Load from SQLite cache first for instant UI response
                    var cached = OfflineCacheService.LoadEmails(node.Id);
                    if (cached != null && cached.Count > 0)
                    {
                        _allEmails.AddRange(cached);
                        _allEmails = _allEmails.OrderByDescending(em => em.Date ?? DateTime.MinValue).ToList();
                        UpdateEmailList();
                    }

                    // 2. Fetch fresh emails asynchronously from Office 365
                    try
                    {
                        var o365Emails = await Office365Service.GetEmailsAsync(node.Id);
                        if (o365Emails != null)
                        {
                            OfflineCacheService.SaveEmails(o365Emails, node.Id);
                            _allEmails.Clear();
                            _allEmails.AddRange(o365Emails);
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Failed to fetch online emails, using SQLite cache: {ex.Message}");
                    }
                }
                else if (node.Folder != null)
                {
                    if (node.Folder.Parent == null)
                    {
                        // This is the root node of the PST file (e.g. context holder)
                        TxtStatus.Text = $"Tệp dữ liệu: {node.Name}";
                        return;
                    }

                    var messages = node.Folder.Messages ?? node.Folder.GetMessages();
                    if (messages != null)
                    {
                        foreach (var msg in messages)
                        {
                            _allEmails.Add(PstParser.MapMessageSummary(msg, node.PstFilePath));
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

    // Xử lý sự kiện Tìm kiếm thời gian thực (Debounced)
    private void TxtSearch_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (TxtSearch == null || TxtSearchPlaceholder == null || BtnClearSearch == null) return;

        _currentSearchText = TxtSearch.Text;
        TxtSearchPlaceholder.Visibility = string.IsNullOrEmpty(_currentSearchText) ? Visibility.Visible : Visibility.Collapsed;
        BtnClearSearch.Visibility = string.IsNullOrEmpty(_currentSearchText) ? Visibility.Collapsed : Visibility.Visible;

        if (_searchDebounceTimer == null)
        {
            _searchDebounceTimer = new System.Windows.Threading.DispatcherTimer();
            _searchDebounceTimer.Interval = TimeSpan.FromMilliseconds(300);
            _searchDebounceTimer.Tick += (s, ev) =>
            {
                _searchDebounceTimer.Stop();
                UpdateEmailList();
            };
        }

        _searchDebounceTimer.Stop();
        _searchDebounceTimer.Start();
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

    private void BtnSort_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button btn)
        {
            btn.ContextMenu.IsOpen = true;
        }
    }

    private void MnuSort_Click(object sender, RoutedEventArgs e)
    {
        if (sender is MenuItem clickedItem)
        {
            foreach (var item in MnuSort.Items)
            {
                if (item is MenuItem mi)
                {
                    mi.IsChecked = (mi == clickedItem);
                }
            }

            _currentSortTag = clickedItem.Tag as string ?? "DateDesc";
            UpdateEmailList();
        }
    }



    private void ChkSelectEmail_Checked(object sender, RoutedEventArgs e)
    {
        UpdateBulkCommandBarVisibility();
    }

    private void ChkSelectEmail_Unchecked(object sender, RoutedEventArgs e)
    {
        UpdateBulkCommandBarVisibility();
    }

    private void UpdateBulkCommandBarVisibility()
    {
        // Bulk command bar UI elements were removed
    }

    private void BtnClearSelection_Click(object sender, RoutedEventArgs e)
    {
        if (_allEmails == null) return;

        foreach (var email in _allEmails)
        {
            email.IsSelected = false;
        }

        UpdateEmailList();
        UpdateBulkCommandBarVisibility();
    }

    private async void BtnBulkDelete_Click(object sender, RoutedEventArgs e)
    {
        if (_allEmails == null) return;

        var selected = _allEmails.Where(x => x.IsSelected).ToList();
        if (!selected.Any()) return;

        if (MessageBox.Show($"Bạn có chắc chắn muốn xóa {selected.Count} email đã chọn không?", "Xóa hàng loạt", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
        {
            try
            {
                TxtStatus.Text = $"Đang xóa {selected.Count} email...";
                foreach (var email in selected)
                {
                    if (email.FilePath.StartsWith("o365://"))
                    {
                        string messageId = email.FilePath.Substring("o365://".Length);
                        await Office365Service.DeleteEmailAsync(messageId);
                    }
                    _allEmails.Remove(email);
                }

                MessageBox.Show($"Đã xóa thành công {selected.Count} email!", "Thành công", MessageBoxButton.OK, MessageBoxImage.Information);
                TxtStatus.Text = "Đã hoàn tất xóa hàng loạt.";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi xóa hàng loạt: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                BtnClearSelection_Click(sender, e);
            }
        }
    }

    private async void BtnBulkArchive_Click(object sender, RoutedEventArgs e)
    {
        if (_allEmails == null) return;

        var selected = _allEmails.Where(x => x.IsSelected).ToList();
        if (!selected.Any()) return;

        try
        {
            TxtStatus.Text = $"Đang lưu trữ {selected.Count} email...";
            foreach (var email in selected)
            {
                if (email.FilePath.StartsWith("o365://"))
                {
                    string messageId = email.FilePath.Substring("o365://".Length);
                    await Office365Service.ArchiveEmailAsync(messageId);
                }
                _allEmails.Remove(email);
            }

            MessageBox.Show($"Đã lưu trữ thành công {selected.Count} email!", "Thành công", MessageBoxButton.OK, MessageBoxImage.Information);
            TxtStatus.Text = "Đã hoàn tất lưu trữ hàng loạt.";
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Lỗi khi lưu trữ hàng loạt: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            BtnClearSelection_Click(sender, e);
        }
    }

    private async void BtnBulkMarkRead_Click(object sender, RoutedEventArgs e)
    {
        if (_allEmails == null) return;

        var selected = _allEmails.Where(x => x.IsSelected).ToList();
        if (!selected.Any()) return;

        try
        {
            TxtStatus.Text = $"Đang đánh dấu đã đọc cho {selected.Count} email...";
            foreach (var email in selected)
            {
                if (email.FilePath.StartsWith("o365://"))
                {
                    string messageId = email.FilePath.Substring("o365://".Length);
                    await Office365Service.MarkAsReadAsync(messageId, true);
                }
                email.IsRead = true;
            }

            TxtStatus.Text = "Đã đánh dấu đã đọc thành công.";
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Lỗi đánh dấu đã đọc hàng loạt: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            BtnClearSelection_Click(sender, e);
        }
    }

    private async void BtnBulkMarkUnread_Click(object sender, RoutedEventArgs e)
    {
        if (_allEmails == null) return;

        var selected = _allEmails.Where(x => x.IsSelected).ToList();
        if (!selected.Any()) return;

        try
        {
            TxtStatus.Text = $"Đang đánh dấu chưa đọc cho {selected.Count} email...";
            foreach (var email in selected)
            {
                if (email.FilePath.StartsWith("o365://"))
                {
                    string messageId = email.FilePath.Substring("o365://".Length);
                    await Office365Service.MarkAsReadAsync(messageId, false);
                }
                email.IsRead = false;
            }

            TxtStatus.Text = "Đã đánh dấu chưa đọc thành công.";
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Lỗi đánh dấu chưa đọc hàng loạt: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            BtnClearSelection_Click(sender, e);
        }
    }

    private void BtnToggleFilters_Click(object sender, RoutedEventArgs e)
    {
        if (PanelAdvancedFilters != null)
        {
            if (PanelAdvancedFilters.Visibility == Visibility.Visible)
            {
                BtnToggleFilters_Unchecked(sender, e);
            }
            else
            {
                BtnToggleFilters_Checked(sender, e);
            }
        }
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
            if (ChkUnreadOnly != null) ChkUnreadOnly.IsChecked = false;
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

    private EmailMessage? GetEmailFromSender(object sender)
    {
        return (sender as MenuItem)?.DataContext as EmailMessage;
    }

    private void MenuReply_Click(object sender, RoutedEventArgs e)
    {
        var email = GetEmailFromSender(sender);
        if (email == null || !Office365Service.IsSignedIn) return;
        var composeWin = new ComposeWindow(email, "reply") { Owner = this };
        composeWin.ShowDialog();
    }

    private void MenuReplyAll_Click(object sender, RoutedEventArgs e)
    {
        var email = GetEmailFromSender(sender);
        if (email == null || !Office365Service.IsSignedIn) return;
        var composeWin = new ComposeWindow(email, "replyall") { Owner = this };
        composeWin.ShowDialog();
    }

    private void MenuForward_Click(object sender, RoutedEventArgs e)
    {
        var email = GetEmailFromSender(sender);
        if (email == null || !Office365Service.IsSignedIn) return;
        var composeWin = new ComposeWindow(email, "forward") { Owner = this };
        composeWin.ShowDialog();
    }

    private async void MenuMarkRead_Click(object sender, RoutedEventArgs e)
    {
        var email = GetEmailFromSender(sender);
        if (email == null || !email.FilePath.StartsWith("o365://")) return;

        try
        {
            string messageId = email.FilePath.Substring("o365://".Length);
            bool nextState = !email.IsRead;
            TxtStatus.Text = nextState ? "Đang đánh dấu đã đọc..." : "Đang đánh dấu chưa đọc...";
            await Office365Service.MarkAsReadAsync(messageId, nextState);
            email.IsRead = nextState;
            LstEmails.Items.Refresh();
            TxtStatus.Text = nextState ? "Đã đánh dấu là Đã đọc." : "Đã đánh dấu là Chưa đọc.";
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Lỗi thay đổi trạng thái đọc: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async void MenuArchive_Click(object sender, RoutedEventArgs e)
    {
        var email = GetEmailFromSender(sender);
        if (email == null || !email.FilePath.StartsWith("o365://")) return;

        try
        {
            TxtStatus.Text = "Đang lưu trữ thư...";
            string messageId = email.FilePath.Substring("o365://".Length);
            await Office365Service.ArchiveEmailAsync(messageId);
            
            _allEmails.Remove(email);
            UpdateEmailList();
            
            if (_currentEmail == email)
            {
                GridPlaceholder.Visibility = Visibility.Visible;
                GridDetail.Visibility = Visibility.Collapsed;
                _currentEmail = null;
            }
            
            TxtStatus.Text = $"Đã lưu trữ email: {email.Subject}";
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Lỗi lưu trữ thư: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async void MenuDelete_Click(object sender, RoutedEventArgs e)
    {
        var email = GetEmailFromSender(sender);
        if (email == null) return;

        if (MessageBox.Show("Bạn có chắc chắn muốn xóa email này không?", "Xác nhận xóa", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
        {
            try
            {
                if (email.FilePath.StartsWith("o365://"))
                {
                    TxtStatus.Text = "Đang xóa thư...";
                    string messageId = email.FilePath.Substring("o365://".Length);
                    await Office365Service.DeleteEmailAsync(messageId);
                }

                _allEmails.Remove(email);
                UpdateEmailList();

                if (_currentEmail == email)
                {
                    GridPlaceholder.Visibility = Visibility.Visible;
                    GridDetail.Visibility = Visibility.Collapsed;
                    _currentEmail = null;
                }

                TxtStatus.Text = $"Đã xóa email: {email.Subject}";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi xóa thư: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }

    private async void MenuJunk_Click(object sender, RoutedEventArgs e)
    {
        var email = GetEmailFromSender(sender);
        if (email == null || !email.FilePath.StartsWith("o365://")) return;

        try
        {
            TxtStatus.Text = "Đang di chuyển vào thư mục Junk Email...";
            string messageId = email.FilePath.Substring("o365://".Length);
            await Office365Service.MoveToJunkAsync(messageId);
            
            _allEmails.Remove(email);
            UpdateEmailList();
            
            if (_currentEmail == email)
            {
                GridPlaceholder.Visibility = Visibility.Visible;
                GridDetail.Visibility = Visibility.Collapsed;
                _currentEmail = null;
            }
            
            TxtStatus.Text = $"Đã di chuyển email sang Junk Email: {email.Subject}";
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Lỗi báo cáo thư rác: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void StartO365Polling()
    {
        if (_o365NewMailTimer == null)
        {
            _o365NewMailTimer = new System.Windows.Threading.DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(30)
            };
            _o365NewMailTimer.Tick += O365NewMailTimer_Tick;
        }
        _o365NewMailTimer.Start();
    }

    private void StopO365Polling()
    {
        if (_o365NewMailTimer != null)
        {
            _o365NewMailTimer.Stop();
        }
    }

    private async void O365NewMailTimer_Tick(object? sender, EventArgs e)
    {
        if (!Office365Service.IsSignedIn)
        {
            StopO365Polling();
            return;
        }

        try
        {
            var latestEmails = await Office365Service.GetEmailsAsync("Inbox");
            if (latestEmails == null || !latestEmails.Any()) return;

            bool hasNew = false;
            foreach (var email in latestEmails)
            {
                if (!_allEmails.Any(em => em.FilePath == email.FilePath))
                {
                    hasNew = true;
                    _allEmails.Insert(0, email);
                    ShowNewMailNotification(email);
                }
            }

            if (hasNew)
            {
                _allEmails = _allEmails.OrderByDescending(em => em.Date ?? DateTime.MinValue).ToList();
                UpdateEmailList();
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error polling new email: {ex.Message}");
        }
    }

    private void ShowNewMailNotification(EmailMessage email)
    {
        var notifyWin = new Window
        {
            Title = "Thư mới",
            Width = 320,
            Height = 85,
            WindowStyle = WindowStyle.None,
            AllowsTransparency = true,
            Background = System.Windows.Media.Brushes.Transparent,
            ShowInTaskbar = false,
            Topmost = true,
            SizeToContent = SizeToContent.WidthAndHeight
        };

        var border = new Border
        {
            Background = new System.Windows.Media.SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#ffffff")),
            BorderBrush = new System.Windows.Media.SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#0078d4")),
            BorderThickness = new Thickness(4, 1, 1, 1),
            CornerRadius = new CornerRadius(0, 4, 4, 0),
            Padding = new Thickness(14, 10, 14, 10),
            Width = 320,
            Effect = new System.Windows.Media.Effects.DropShadowEffect { Color = System.Windows.Media.Colors.Gray, BlurRadius = 10, ShadowDepth = 2, Opacity = 0.3 }
        };

        var grid = new Grid();
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        var title = new TextBlock { Text = $"📩 Thư mới từ {email.FromName}", FontWeight = FontWeights.Bold, FontSize = 12, Foreground = System.Windows.Media.Brushes.Black, TextTrimming = TextTrimming.CharacterEllipsis };
        var subject = new TextBlock { Text = email.Subject, FontSize = 11, Foreground = System.Windows.Media.Brushes.Gray, Margin = new Thickness(0, 4, 0, 0), TextTrimming = TextTrimming.CharacterEllipsis };

        Grid.SetRow(title, 0);
        Grid.SetRow(subject, 1);
        grid.Children.Add(title);
        grid.Children.Add(subject);
        border.Child = grid;
        notifyWin.Content = border;

        var workingArea = SystemParameters.WorkArea;
        notifyWin.Left = workingArea.Right - notifyWin.Width - 10;
        notifyWin.Top = workingArea.Bottom - notifyWin.Height - 10;

        notifyWin.Loaded += async (s, e) =>
        {
            await Task.Delay(4000);
            notifyWin.Close();
        };

        notifyWin.Show();
    }

    private static bool IsMatch(string? target, string keyword)
    {
        if (string.IsNullOrEmpty(target)) return false;
        if (string.IsNullOrEmpty(keyword)) return false;

        // Normalize both to Form C
        string normTarget = target.Normalize(System.Text.NormalizationForm.FormC);
        string normKeyword = keyword.Normalize(System.Text.NormalizationForm.FormC);

        var keywordTerms = normKeyword.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
        if (keywordTerms.Length == 0) return false;

        var targetWords = SplitIntoWords(normTarget);

        foreach (var term in keywordTerms)
        {
            bool termMatched = false;
            string cleanTerm = RemoveDiacritics(term);
            bool isTermAccented = (cleanTerm != term);

            foreach (var word in targetWords)
            {
                if (isTermAccented)
                {
                    if (word.StartsWith(term, StringComparison.OrdinalIgnoreCase))
                    {
                        termMatched = true;
                        break;
                    }
                }
                else
                {
                    string cleanWord = RemoveDiacritics(word);
                    if (cleanWord.StartsWith(term, StringComparison.OrdinalIgnoreCase))
                    {
                        termMatched = true;
                        break;
                    }
                }
            }

            if (!termMatched) return false;
        }

        return true;
    }

    private static List<string> SplitIntoWords(string text)
    {
        var words = new List<string>();
        if (string.IsNullOrEmpty(text)) return words;

        var sb = new System.Text.StringBuilder();
        foreach (var c in text)
        {
            if (char.IsLetterOrDigit(c))
            {
                sb.Append(c);
            }
            else
            {
                if (sb.Length > 0)
                {
                    words.Add(sb.ToString());
                    sb.Clear();
                }
            }
        }
        if (sb.Length > 0)
        {
            words.Add(sb.ToString());
        }
        return words;
    }

    private static string RemoveDiacritics(string text)
    {
        if (string.IsNullOrEmpty(text)) return string.Empty;
        var normalizedString = text.Normalize(System.Text.NormalizationForm.FormD);
        var stringBuilder = new System.Text.StringBuilder();

        foreach (var c in normalizedString)
        {
            var unicodeCategory = System.Globalization.CharUnicodeInfo.GetUnicodeCategory(c);
            if (unicodeCategory != System.Globalization.UnicodeCategory.NonSpacingMark)
            {
                if (c == 'đ') stringBuilder.Append('d');
                else if (c == 'Đ') stringBuilder.Append('D');
                else stringBuilder.Append(c);
            }
        }

        return stringBuilder.ToString().Normalize(System.Text.NormalizationForm.FormC);
    }

    private static void ParseOutlookSearchQuery(string query, out string? subject, out string? from, out string? to, out bool? hasAttachment, out bool? isUnread, out string? generalKeyword)
    {
        subject = null;
        from = null;
        to = null;
        hasAttachment = null;
        isUnread = null;
        generalKeyword = null;

        if (string.IsNullOrWhiteSpace(query)) return;

        // Split query by spaces, but respect quotes for phrases
        var terms = System.Text.RegularExpressions.Regex.Matches(query, @"(?<match>\w+:""[^""]+""|\w+:[^\s]+|""[^""]+""|[^\s]+)");
        var generalList = new List<string>();

        foreach (System.Text.RegularExpressions.Match termMatch in terms)
        {
            string term = termMatch.Value;
            if (term.Contains(":"))
            {
                int colonIndex = term.IndexOf(':');
                string key = term.Substring(0, colonIndex).ToLowerInvariant();
                string val = term.Substring(colonIndex + 1).Trim('"', '\'');

                switch (key)
                {
                    case "from":
                        from = val;
                        break;
                    case "subject":
                        subject = val;
                        break;
                    case "to":
                        to = val;
                        break;
                    case "has":
                    case "hasattachment":
                        if (val.Equals("attachment", StringComparison.OrdinalIgnoreCase) || val.Equals("yes", StringComparison.OrdinalIgnoreCase) || val.Equals("true", StringComparison.OrdinalIgnoreCase))
                            hasAttachment = true;
                        else if (val.Equals("no", StringComparison.OrdinalIgnoreCase) || val.Equals("false", StringComparison.OrdinalIgnoreCase))
                            hasAttachment = false;
                        break;
                    case "is":
                        if (val.Equals("unread", StringComparison.OrdinalIgnoreCase))
                            isUnread = true;
                        else if (val.Equals("read", StringComparison.OrdinalIgnoreCase))
                            isUnread = false;
                        break;
                    default:
                        generalList.Add(term);
                        break;
                }
            }
            else
            {
                generalList.Add(term.Trim('"', '\''));
            }
        }

        if (generalList.Count > 0)
        {
            generalKeyword = string.Join(" ", generalList);
        }
    }

    protected override void OnClosed(EventArgs e)
    {
        foreach (var pst in _activePstFiles.Values)
        {
            try { pst.Dispose(); } catch {}
        }
        _activePstFiles.Clear();
        StopO365Polling();
        base.OnClosed(e);
    }
}

public static class TextBlockHelper
{
    public static readonly DependencyProperty HighlightTextProperty =
        DependencyProperty.RegisterAttached(
            "HighlightText",
            typeof(string),
            typeof(TextBlockHelper),
            new PropertyMetadata(null, OnHighlightChanged));

    public static readonly DependencyProperty KeywordProperty =
        DependencyProperty.RegisterAttached(
            "Keyword",
            typeof(string),
            typeof(TextBlockHelper),
            new PropertyMetadata(null, OnHighlightChanged));

    public static string GetHighlightText(DependencyObject obj) => (string)obj.GetValue(HighlightTextProperty);
    public static void SetHighlightText(DependencyObject obj, string value) => obj.SetValue(HighlightTextProperty, value);

    public static string GetKeyword(DependencyObject obj) => (string)obj.GetValue(KeywordProperty);
    public static void SetKeyword(DependencyObject obj, string value) => obj.SetValue(KeywordProperty, value);

    private static void OnHighlightChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        if (d is TextBlock textBlock)
        {
            string text = GetHighlightText(textBlock) ?? "";
            string keyword = GetKeyword(textBlock) ?? "";

            textBlock.Inlines.Clear();

            if (string.IsNullOrEmpty(text)) return;

            if (string.IsNullOrEmpty(keyword))
            {
                textBlock.Inlines.Add(new System.Windows.Documents.Run(text));
                return;
            }

            string normText = text.Normalize(System.Text.NormalizationForm.FormC);
            string normKeyword = keyword.Normalize(System.Text.NormalizationForm.FormC);

            var terms = normKeyword.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)
                                   .Select(t => RemoveDiacritics(t).ToLowerInvariant())
                                   .Where(t => !string.IsNullOrEmpty(t))
                                   .ToList();

            if (terms.Count == 0)
            {
                textBlock.Inlines.Add(new System.Windows.Documents.Run(text));
                return;
            }

            int index = 0;
            while (index < text.Length)
            {
                // Add non-word characters
                while (index < text.Length && !char.IsLetterOrDigit(text[index]))
                {
                    textBlock.Inlines.Add(new System.Windows.Documents.Run(text[index].ToString()));
                    index++;
                }

                if (index >= text.Length) break;

                // Extract word
                int wordStart = index;
                while (index < text.Length && char.IsLetterOrDigit(text[index]))
                {
                    index++;
                }
                string word = text.Substring(wordStart, index - wordStart);
                string cleanWord = RemoveDiacritics(word).ToLowerInvariant();

                string? matchedTerm = null;
                foreach (var term in terms)
                {
                    if (cleanWord.StartsWith(term))
                    {
                        matchedTerm = term;
                        break;
                    }
                }

                if (matchedTerm != null && word.Length >= matchedTerm.Length)
                {
                    int highlightLength = matchedTerm.Length;
                    int originalHighlightLength = 0;
                    for (int len = 1; len <= word.Length; len++)
                    {
                        string sub = word.Substring(0, len);
                        if (RemoveDiacritics(sub).ToLowerInvariant().Length >= highlightLength)
                        {
                            originalHighlightLength = len;
                            break;
                        }
                    }
                    if (originalHighlightLength == 0) originalHighlightLength = highlightLength;

                    string prefix = word.Substring(0, originalHighlightLength);
                    string suffix = word.Substring(originalHighlightLength);

                    var highlightRun = new System.Windows.Documents.Run(prefix)
                    {
                        Background = new System.Windows.Media.SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#ffeb3b")),
                        FontWeight = FontWeights.Bold
                    };
                    textBlock.Inlines.Add(highlightRun);

                    if (!string.IsNullOrEmpty(suffix))
                    {
                        textBlock.Inlines.Add(new System.Windows.Documents.Run(suffix));
                    }
                }
                else
                {
                    textBlock.Inlines.Add(new System.Windows.Documents.Run(word));
                }
            }
        }
    }

    private static string RemoveDiacritics(string text)
    {
        if (string.IsNullOrEmpty(text)) return string.Empty;
        var normalizedString = text.Normalize(System.Text.NormalizationForm.FormD);
        var stringBuilder = new System.Text.StringBuilder();

        foreach (var c in normalizedString)
        {
            var unicodeCategory = System.Globalization.CharUnicodeInfo.GetUnicodeCategory(c);
            if (unicodeCategory != System.Globalization.UnicodeCategory.NonSpacingMark)
            {
                if (c == 'đ') stringBuilder.Append('d');
                else if (c == 'Đ') stringBuilder.Append('D');
                else stringBuilder.Append(c);
            }
        }

        return stringBuilder.ToString().Normalize(System.Text.NormalizationForm.FormC);
    }
}