using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Win32;
using MsgViewer.Services;

namespace MsgViewer
{
    /// <summary>
    /// Interaction logic for ComposeWindow.xaml
    /// </summary>
    public partial class ComposeWindow : Window
    {
        private readonly List<AttachmentItem> _attachments = new();

        public ComposeWindow()
        {
            InitializeComponent();
            RefreshAttachmentsList();
        }

        private void BtnAttach_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Multiselect = true,
                Title = "Chọn tệp đính kèm"
            };

            if (dialog.ShowDialog() == true)
            {
                foreach (var filePath in dialog.FileNames)
                {
                    if (_attachments.All(a => a.FilePath != filePath))
                    {
                        _attachments.Add(new AttachmentItem
                        {
                            FileName = Path.GetFileName(filePath),
                            FilePath = filePath
                        });
                    }
                }
                RefreshAttachmentsList();
            }
        }

        private void BtnRemoveAttachment_Click(object sender, RoutedEventArgs e)
        {
            if (sender is System.Windows.Controls.Button btn && btn.DataContext is AttachmentItem item)
            {
                _attachments.Remove(item);
                RefreshAttachmentsList();
            }
        }

        private void RefreshAttachmentsList()
        {
            LstAttachments.ItemsSource = null;
            LstAttachments.ItemsSource = _attachments;
            PanelAttachments.Visibility = _attachments.Any() ? Visibility.Visible : Visibility.Collapsed;
        }

        private async void BtnSend_Click(object sender, RoutedEventArgs e)
        {
            string to = TxtTo.Text.Trim();
            string cc = TxtCc.Text.Trim();
            string subject = TxtSubject.Text.Trim();
            string body = TxtBody.Text;

            if (string.IsNullOrWhiteSpace(to))
            {
                MessageBox.Show("Vui lòng nhập địa chỉ email người nhận (To).", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Convert plain text body to simple HTML paragraphs
            string htmlBody = "<html><body>" + 
                string.Join("<br/>", body.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None)
                .Select(line => System.Net.WebUtility.HtmlEncode(line))) + 
                "</body></html>";

            // Disable controls to prevent double submission
            SetControlsEnabled(false);
            TxtStatus.Text = "Đang gửi email qua Office 365...";

            try
            {
                var attachmentsList = _attachments.Select(a => a.FilePath).ToList();
                bool result = await Office365Service.SendEmailAsync(to, cc, subject, htmlBody, attachmentsList);
                if (result)
                {
                    MessageBox.Show("Email đã được gửi thành công!", "Thành công", MessageBoxButton.OK, MessageBoxImage.Information);
                    this.DialogResult = true;
                    this.Close();
                }
                else
                {
                    MessageBox.Show("Không thể gửi email. Vui lòng kiểm tra lại cấu hình kết nối.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                    SetControlsEnabled(true);
                    TxtStatus.Text = "Gửi thư thất bại.";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi gửi email: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                SetControlsEnabled(true);
                TxtStatus.Text = $"Lỗi: {ex.Message}";
            }
        }

        private void SetControlsEnabled(bool enabled)
        {
            BtnSend.IsEnabled = enabled;
            BtnAttach.IsEnabled = enabled;
            TxtTo.IsEnabled = enabled;
            TxtCc.IsEnabled = enabled;
            TxtSubject.IsEnabled = enabled;
            TxtBody.IsEnabled = enabled;
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }

    public class AttachmentItem
    {
        public string FileName { get; set; } = "";
        public string FilePath { get; set; } = "";
    }
}
