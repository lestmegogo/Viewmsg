using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;
using Microsoft.Win32;
using MsgViewer.Models;
using MsgViewer.Services;

namespace MsgViewer
{
    /// <summary>
    /// Interaction logic for ComposeWindow.xaml
    /// </summary>
    public partial class ComposeWindow : Window
    {
        private readonly List<AttachmentItem> _attachments = new();
        private System.Windows.Threading.DispatcherTimer? _autoSaveTimer;
        private string? _draftId;
        private string _lastSavedContent = "";
        private List<(string Name, string Email)> _contacts = new();

        public ComposeWindow() : this(null, "new")
        {
        }

        public ComposeWindow(EmailMessage? sourceEmail, string actionType)
        {
            InitializeComponent();
            RefreshAttachmentsList();

            if (sourceEmail != null)
            {
                if (actionType.Equals("reply", StringComparison.OrdinalIgnoreCase))
                {
                    TxtTo.Text = sourceEmail.FromEmail + "; ";
                    TxtSubject.Text = sourceEmail.Subject.StartsWith("Re:", StringComparison.OrdinalIgnoreCase) 
                        ? sourceEmail.Subject 
                        : "Re: " + sourceEmail.Subject;
                    PopulateQuotedContent(sourceEmail.BodyText, actionType, sourceEmail.FromDisplay, sourceEmail.DateDisplay, sourceEmail.Subject);
                }
                else if (actionType.Equals("replyall", StringComparison.OrdinalIgnoreCase))
                {
                    var recipients = new List<string>();
                    if (!string.IsNullOrEmpty(sourceEmail.FromEmail)) recipients.Add(sourceEmail.FromEmail);
                    if (!string.IsNullOrEmpty(sourceEmail.To)) recipients.Add(sourceEmail.To);
                    
                    TxtTo.Text = string.Join("; ", recipients.Distinct()) + "; ";
                    TxtCc.Text = sourceEmail.Cc;
                    TxtSubject.Text = sourceEmail.Subject.StartsWith("Re:", StringComparison.OrdinalIgnoreCase) 
                        ? sourceEmail.Subject 
                        : "Re: " + sourceEmail.Subject;
                    PopulateQuotedContent(sourceEmail.BodyText, actionType, sourceEmail.FromDisplay, sourceEmail.DateDisplay, sourceEmail.Subject);
                }
                else if (actionType.Equals("forward", StringComparison.OrdinalIgnoreCase))
                {
                    TxtSubject.Text = sourceEmail.Subject.StartsWith("Fw:", StringComparison.OrdinalIgnoreCase) 
                        ? sourceEmail.Subject 
                        : "Fw: " + sourceEmail.Subject;
                    PopulateQuotedContent(sourceEmail.BodyText, actionType, sourceEmail.FromDisplay, sourceEmail.DateDisplay, sourceEmail.Subject);
                    
                    // Copy attachments
                    foreach (var att in sourceEmail.Attachments)
                    {
                        try
                        {
                            string tempPath = Path.Combine(Path.GetTempPath(), att.FileName);
                            File.WriteAllBytes(tempPath, att.Data);
                            _attachments.Add(new AttachmentItem { FileName = att.FileName, FilePath = tempPath });
                        }
                        catch {}
                    }
                    RefreshAttachmentsList();
                }
                else if (actionType.Equals("draft", StringComparison.OrdinalIgnoreCase))
                {
                    _draftId = sourceEmail.FilePath.StartsWith("o365://") 
                        ? sourceEmail.FilePath.Substring("o365://".Length) 
                        : null;
                    TxtTo.Text = sourceEmail.To;
                    TxtCc.Text = sourceEmail.Cc;
                    TxtSubject.Text = sourceEmail.Subject;
                    LoadHtmlIntoFlowDocument(sourceEmail.BodyHtml ?? "");
                }
            }

            InitializeAutoSave();
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
            string htmlBody = ConvertFlowDocumentToHtml(RtfBody.Document);

            if (string.IsNullOrWhiteSpace(to))
            {
                MessageBox.Show("Vui lòng nhập địa chỉ email người nhận (To).", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Disable controls to prevent double submission
            SetControlsEnabled(false);
            TxtStatus.Text = "Đang gửi email qua Office 365...";

            try
            {
                var attachmentsList = _attachments.Select(a => a.FilePath).ToList();
                bool result = await Office365Service.SendEmailAsync(to, cc, subject, htmlBody, attachmentsList);
                if (result)
                {
                    if (_autoSaveTimer != null) _autoSaveTimer.Stop();
                    if (!string.IsNullOrEmpty(_draftId))
                    {
                        try
                        {
                            await Office365Service.DeleteEmailAsync(_draftId);
                        }
                        catch {}
                    }
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
            if (RtfBody != null) RtfBody.IsEnabled = enabled;
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void BtnFormatBold_Click(object sender, RoutedEventArgs e)
        {
            if (RtfBody.Selection != null)
            {
                var isBold = RtfBody.Selection.GetPropertyValue(TextElement.FontWeightProperty);
                if (isBold is FontWeight weight && weight == FontWeights.Bold)
                {
                    RtfBody.Selection.ApplyPropertyValue(TextElement.FontWeightProperty, FontWeights.Normal);
                }
                else
                {
                    RtfBody.Selection.ApplyPropertyValue(TextElement.FontWeightProperty, FontWeights.Bold);
                }
            }
        }

        private void BtnFormatItalic_Click(object sender, RoutedEventArgs e)
        {
            if (RtfBody.Selection != null)
            {
                var isItalic = RtfBody.Selection.GetPropertyValue(TextElement.FontStyleProperty);
                if (isItalic is FontStyle style && style == FontStyles.Italic)
                {
                    RtfBody.Selection.ApplyPropertyValue(TextElement.FontStyleProperty, FontStyles.Normal);
                }
                else
                {
                    RtfBody.Selection.ApplyPropertyValue(TextElement.FontStyleProperty, FontStyles.Italic);
                }
            }
        }

        private void BtnFormatUnderline_Click(object sender, RoutedEventArgs e)
        {
            if (RtfBody.Selection != null)
            {
                var isUnderline = RtfBody.Selection.GetPropertyValue(Inline.TextDecorationsProperty);
                if (isUnderline != null && isUnderline != DependencyProperty.UnsetValue && isUnderline.Equals(TextDecorations.Underline))
                {
                    RtfBody.Selection.ApplyPropertyValue(Inline.TextDecorationsProperty, null);
                }
                else
                {
                    RtfBody.Selection.ApplyPropertyValue(Inline.TextDecorationsProperty, TextDecorations.Underline);
                }
            }
        }

        private void BtnFormatList_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Documents.EditingCommands.ToggleBullets.Execute(null, RtfBody);
        }

        private void BtnFormatColor_Click(object sender, RoutedEventArgs e)
        {
            if (RtfBody.Selection != null)
            {
                var isBlue = RtfBody.Selection.GetPropertyValue(TextElement.ForegroundProperty);
                if (isBlue is System.Windows.Media.SolidColorBrush brush && brush.Color == System.Windows.Media.Colors.Blue)
                {
                    RtfBody.Selection.ApplyPropertyValue(TextElement.ForegroundProperty, System.Windows.Media.Brushes.Black);
                }
                else
                {
                    RtfBody.Selection.ApplyPropertyValue(TextElement.ForegroundProperty, System.Windows.Media.Brushes.Blue);
                }
            }
        }

        private void RtfBody_SelectionChanged(object sender, RoutedEventArgs e)
        {
        }

        private void RtfBody_PreviewDragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
                e.Handled = true;
            }
        }

        private void RtfBody_Drop(object sender, DragEventArgs e)
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
                            var existing = _attachments.FirstOrDefault(a => a.FilePath == file);
                            if (existing == null)
                            {
                                _attachments.Add(new AttachmentItem 
                                { 
                                    FileName = Path.GetFileName(file), 
                                    FilePath = file 
                                });
                            }
                        }
                    }
                    RefreshAttachmentsList();
                }
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
                    string colorHex = fg.Color.ToString().Substring(3); // Remove Alpha channel (e.g. #FF0000FF -> 0000FF)
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

        private void InitializeAutoSave()
        {
            _autoSaveTimer = new System.Windows.Threading.DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(15)
            };
            _autoSaveTimer.Tick += AutoSaveTimer_Tick;
            _autoSaveTimer.Start();
        }

        private async void AutoSaveTimer_Tick(object? sender, EventArgs e)
        {
            await SaveDraftAsync();
        }

        private async Task SaveDraftAsync()
        {
            string to = TxtTo.Text.Trim();
            string cc = TxtCc.Text.Trim();
            string subject = TxtSubject.Text.Trim();
            string htmlBody = ConvertFlowDocumentToHtml(RtfBody.Document);

            // If everything is blank, don't save
            if (string.IsNullOrWhiteSpace(to) && string.IsNullOrWhiteSpace(subject) && (htmlBody == null || htmlBody.Contains("<body></body>") || htmlBody.Contains("<body><p></p></body>")))
            {
                return;
            }

            string currentContent = $"{to}|{cc}|{subject}|{htmlBody}";
            if (currentContent == _lastSavedContent) return; // No change

            try
            {
                TxtStatus.Text = "Đang tự động lưu bản nháp...";
                var newDraftId = await Office365Service.CreateOrUpdateDraftAsync(_draftId, to, cc, subject, htmlBody);
                if (newDraftId != null)
                {
                    _draftId = newDraftId;
                    _lastSavedContent = currentContent;
                    TxtStatus.Text = $"Đã tự động lưu nháp lúc {DateTime.Now:HH:mm:ss}";
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Lỗi lưu nháp: {ex.Message}");
            }
        }

        protected override void OnClosed(EventArgs e)
        {
            if (_autoSaveTimer != null)
            {
                _autoSaveTimer.Stop();
                _autoSaveTimer = null;
            }
            base.OnClosed(e);
        }

        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                var list = await Office365Service.GetContactsAsync();
                if (list != null)
                {
                    _contacts = list;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to prefetch contacts: {ex.Message}");
            }
        }

        private void TxtTo_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            if (_contacts == null || _contacts.Count == 0) return;

            string fullText = TxtTo.Text;
            if (string.IsNullOrWhiteSpace(fullText))
            {
                PopupSuggestions.IsOpen = false;
                return;
            }

            // Find current typing segment
            string currentWord = fullText.Split(new[] { ';', ',' }).Last().Trim();
            if (string.IsNullOrEmpty(currentWord))
            {
                PopupSuggestions.IsOpen = false;
                return;
            }

            // Filter contacts
            var filtered = _contacts.Where(c => 
                c.Name.Contains(currentWord, StringComparison.OrdinalIgnoreCase) || 
                c.Email.Contains(currentWord, StringComparison.OrdinalIgnoreCase))
                .Select(c => new { c.Name, c.Email })
                .ToList();

            if (filtered.Any())
            {
                LstSuggestions.ItemsSource = filtered;
                PopupSuggestions.IsOpen = true;
            }
            else
            {
                PopupSuggestions.IsOpen = false;
            }
        }

        private void TxtTo_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (PopupSuggestions.IsOpen)
            {
                if (e.Key == System.Windows.Input.Key.Down)
                {
                    int nextIndex = LstSuggestions.SelectedIndex + 1;
                    if (nextIndex < LstSuggestions.Items.Count)
                    {
                        LstSuggestions.SelectedIndex = nextIndex;
                        LstSuggestions.ScrollIntoView(LstSuggestions.SelectedItem);
                    }
                    e.Handled = true;
                }
                else if (e.Key == System.Windows.Input.Key.Up)
                {
                    int prevIndex = LstSuggestions.SelectedIndex - 1;
                    if (prevIndex >= 0)
                    {
                        LstSuggestions.SelectedIndex = prevIndex;
                        LstSuggestions.ScrollIntoView(LstSuggestions.SelectedItem);
                    }
                    e.Handled = true;
                }
                else if (e.Key == System.Windows.Input.Key.Enter || e.Key == System.Windows.Input.Key.Tab)
                {
                    SelectCurrentSuggestion();
                    e.Handled = true;
                }
            }
        }

        private void LstSuggestions_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            SelectCurrentSuggestion();
        }

        private void SelectCurrentSuggestion()
        {
            var selected = LstSuggestions.SelectedItem;
            if (selected == null) return;

            dynamic item = selected;
            string selectedEmail = item.Email;

            string fullText = TxtTo.Text;
            var parts = fullText.Split(new[] { ';', ',' }).ToList();
            if (parts.Count > 0)
            {
                parts[parts.Count - 1] = $" {selectedEmail}";
            }
            else
            {
                parts.Add(selectedEmail);
            }

            TxtTo.TextChanged -= TxtTo_TextChanged; // Unhook to prevent recursive loops
            TxtTo.Text = string.Join(";", parts).Trim() + "; ";
            TxtTo.CaretIndex = TxtTo.Text.Length;
            TxtTo.TextChanged += TxtTo_TextChanged;

            PopupSuggestions.IsOpen = false;
        }

        private void PopulateQuotedContent(string? bodyText, string actionType, string sender, string dateStr, string subject)
        {
            var doc = RtfBody.Document;
            doc.Blocks.Clear();

            // 1. Add empty paragraph at the top for user typing
            doc.Blocks.Add(new Paragraph(new Run("")));

            // 2. Add divider line / header
            var headerPara = new Paragraph();
            headerPara.Foreground = Brushes.Gray;
            headerPara.Margin = new Thickness(0, 20, 0, 10);
            headerPara.Inlines.Add(new LineBreak());
            headerPara.Inlines.Add(new Run("________________________________________"));
            headerPara.Inlines.Add(new LineBreak());
            headerPara.Inlines.Add(new Run($"Từ (From): {sender}"));
            headerPara.Inlines.Add(new LineBreak());
            headerPara.Inlines.Add(new Run($"Gửi lúc (Sent): {dateStr}"));
            headerPara.Inlines.Add(new LineBreak());
            headerPara.Inlines.Add(new Run($"Tiêu đề (Subject): {subject}"));
            headerPara.Inlines.Add(new LineBreak());
            doc.Blocks.Add(headerPara);

            // 3. Add quoted body text in a BlockQuote style
            if (!string.IsNullOrWhiteSpace(bodyText))
            {
                var bodyPara = new Paragraph();
                bodyPara.Foreground = Brushes.DimGray;
                bodyPara.Margin = new Thickness(20, 0, 0, 0); // Indent

                var lines = bodyText.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
                foreach (var line in lines)
                {
                    bodyPara.Inlines.Add(new Run(line));
                    bodyPara.Inlines.Add(new LineBreak());
                }
                doc.Blocks.Add(bodyPara);
            }
        }

        private void LoadHtmlIntoFlowDocument(string html)
        {
            var doc = RtfBody.Document;
            doc.Blocks.Clear();

            if (string.IsNullOrWhiteSpace(html))
            {
                doc.Blocks.Add(new Paragraph(new Run("")));
                return;
            }

            try
            {
                int bodyStart = html.IndexOf("<body>", StringComparison.OrdinalIgnoreCase);
                if (bodyStart >= 0)
                {
                    int bodyEnd = html.IndexOf("</body>", StringComparison.OrdinalIgnoreCase);
                    if (bodyEnd > bodyStart)
                    {
                        html = html.Substring(bodyStart + 6, bodyEnd - bodyStart - 6);
                    }
                }

                // Split by paragraphs
                var paraSplits = html.Split(new[] { "<p>", "</p>" }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var pText in paraSplits)
                {
                    if (string.IsNullOrWhiteSpace(pText)) continue;

                    var p = new Paragraph();
                    string temp = pText.Replace("<br/>", "\n").Replace("<br>", "\n");
                    ParseParagraphContent(p, temp);
                    doc.Blocks.Add(p);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error parsing draft HTML: {ex.Message}");
                doc.Blocks.Add(new Paragraph(new Run(html)));
            }

            if (!doc.Blocks.Any())
            {
                doc.Blocks.Add(new Paragraph(new Run("")));
            }
        }

        private void ParseParagraphContent(Paragraph p, string text)
        {
            int i = 0;
            while (i < text.Length)
            {
                int nextTag = text.IndexOf('<', i);
                if (nextTag < 0)
                {
                    string remaining = System.Net.WebUtility.HtmlDecode(text.Substring(i));
                    p.Inlines.Add(new Run(remaining));
                    break;
                }

                if (nextTag > i)
                {
                    string prevText = System.Net.WebUtility.HtmlDecode(text.Substring(i, nextTag - i));
                    p.Inlines.Add(new Run(prevText));
                }

                int tagEnd = text.IndexOf('>', nextTag);
                if (tagEnd < 0)
                {
                    p.Inlines.Add(new Run(text.Substring(nextTag)));
                    break;
                }

                string tag = text.Substring(nextTag, tagEnd - nextTag + 1);
                i = tagEnd + 1;

                if (tag.StartsWith("<strong>", StringComparison.OrdinalIgnoreCase))
                {
                    int closeIdx = text.IndexOf("</strong>", i, StringComparison.OrdinalIgnoreCase);
                    if (closeIdx >= 0)
                    {
                        string runText = System.Net.WebUtility.HtmlDecode(text.Substring(i, closeIdx - i));
                        p.Inlines.Add(new Run(runText) { FontWeight = FontWeights.Bold });
                        i = closeIdx + 9;
                    }
                }
                else if (tag.StartsWith("<em>", StringComparison.OrdinalIgnoreCase))
                {
                    int closeIdx = text.IndexOf("</em>", i, StringComparison.OrdinalIgnoreCase);
                    if (closeIdx >= 0)
                    {
                        string runText = System.Net.WebUtility.HtmlDecode(text.Substring(i, closeIdx - i));
                        p.Inlines.Add(new Run(runText) { FontStyle = FontStyles.Italic });
                        i = closeIdx + 5;
                    }
                }
                else if (tag.StartsWith("<u>", StringComparison.OrdinalIgnoreCase))
                {
                    int closeIdx = text.IndexOf("</u>", i, StringComparison.OrdinalIgnoreCase);
                    if (closeIdx >= 0)
                    {
                        string runText = System.Net.WebUtility.HtmlDecode(text.Substring(i, closeIdx - i));
                        var r = new Run(runText);
                        r.TextDecorations = TextDecorations.Underline;
                        p.Inlines.Add(r);
                        i = closeIdx + 4;
                    }
                }
                else if (tag.StartsWith("<span", StringComparison.OrdinalIgnoreCase))
                {
                    int closeIdx = text.IndexOf("</span>", i, StringComparison.OrdinalIgnoreCase);
                    if (closeIdx >= 0)
                    {
                        string runText = System.Net.WebUtility.HtmlDecode(text.Substring(i, closeIdx - i));
                        var r = new Run(runText);
                        
                        int styleStart = tag.IndexOf("style='color:#", StringComparison.OrdinalIgnoreCase);
                        if (styleStart >= 0)
                        {
                            string hex = tag.Substring(styleStart + 14, 6);
                            try
                            {
                                var color = (Color)ColorConverter.ConvertFromString($"#{hex}");
                                r.Foreground = new SolidColorBrush(color);
                            }
                            catch {}
                        }
                        p.Inlines.Add(r);
                        i = closeIdx + 7;
                    }
                }
            }
        }
    }

    public class AttachmentItem
    {
        public string FileName { get; set; } = "";
        public string FilePath { get; set; } = "";
    }
}
