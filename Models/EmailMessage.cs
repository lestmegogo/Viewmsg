using System.Collections.Generic;

namespace MsgViewer.Models;

/// <summary>
/// Mô hình email thống nhất cho cả file .msg và .eml.
/// </summary>
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Collections.ObjectModel;

public class EmailMessage : INotifyPropertyChanged
{
    private string _filePath = "";
    public string FilePath
    {
        get => _filePath;
        set { _filePath = value; OnPropertyChanged(); }
    }

    private string _subject = "(Không có tiêu đề)";
    public string Subject
    {
        get => _subject;
        set { _subject = value; OnPropertyChanged(); }
    }

    private string _fromName = "";
    public string FromName
    {
        get => _fromName;
        set 
        { 
            _fromName = value; 
            OnPropertyChanged(); 
            OnPropertyChanged(nameof(FromDisplay)); 
            OnPropertyChanged(nameof(FromClean)); 
            OnPropertyChanged(nameof(SenderInitials)); 
        }
    }

    private string _fromEmail = "";
    public string FromEmail
    {
        get => _fromEmail;
        set 
        { 
            _fromEmail = value; 
            OnPropertyChanged(); 
            OnPropertyChanged(nameof(FromDisplay)); 
            OnPropertyChanged(nameof(FromClean)); 
            OnPropertyChanged(nameof(SenderInitials)); 
        }
    }

    private string _to = "";
    public string To
    {
        get => _to;
        set { _to = value; OnPropertyChanged(); }
    }

    private string _cc = "";
    public string Cc
    {
        get => _cc;
        set { _cc = value; OnPropertyChanged(); }
    }

    private DateTime? _date;
    public DateTime? Date
    {
        get => _date;
        set 
        { 
            _date = value; 
            OnPropertyChanged(); 
            OnPropertyChanged(nameof(DateGroup)); 
            OnPropertyChanged(nameof(DateDisplay)); 
        }
    }

    private bool _isRead = true;
    public bool IsRead
    {
        get => _isRead;
        set { _isRead = value; OnPropertyChanged(); }
    }

    private bool _isSelected = false;
    public bool IsSelected
    {
        get => _isSelected;
        set { _isSelected = value; OnPropertyChanged(); }
    }

    public string DateGroup
    {
        get
        {
            if (!Date.HasValue) return "Cũ hơn (Older)";
            var date = Date.Value.Date;
            if (date == DateTime.Today) return "Hôm nay (Today)";
            if (date == DateTime.Today.AddDays(-1)) return "Hôm qua (Yesterday)";
            if (date >= DateTime.Today.AddDays(-7)) return "Tuần này (This week)";
            if (date >= DateTime.Today.AddDays(-30)) return "Tháng này (This month)";
            return "Cũ hơn (Older)";
        }
    }

    private string? _bodyHtml;
    public string? BodyHtml
    {
        get => _bodyHtml;
        set { _bodyHtml = value; OnPropertyChanged(); }
    }

    private string? _bodyText;
    public string? BodyText
    {
        get => _bodyText;
        set 
        { 
            _bodyText = value; 
            OnPropertyChanged(); 
            OnPropertyChanged(nameof(Snippet)); 
        }
    }

    public ObservableCollection<EmailAttachment> Attachments { get; } = new();

    public string FromDisplay =>
        string.IsNullOrWhiteSpace(FromName)
            ? FromEmail
            : string.IsNullOrWhiteSpace(FromEmail) ? FromName : $"{FromName} <{FromEmail}>";

    public string FromClean => string.IsNullOrWhiteSpace(FromName) ? FromEmail : FromName;

    public string DateDisplay => Date?.ToString("dddd, dd/MM/yyyy HH:mm") ?? "";

    private string _snippet = "";
    public string Snippet
    {
        get
        {
            if (!string.IsNullOrWhiteSpace(_snippet)) return _snippet;

            string source = "";
            if (!string.IsNullOrWhiteSpace(BodyText))
            {
                source = BodyText;
            }
            else if (!string.IsNullOrWhiteSpace(BodyHtml))
            {
                source = System.Text.RegularExpressions.Regex.Replace(BodyHtml, "<.*?>", string.Empty);
                source = System.Net.WebUtility.HtmlDecode(source);
            }

            if (string.IsNullOrWhiteSpace(source)) return "";

            var text = source.Replace("\r", " ").Replace("\n", " ").Trim();
            text = System.Text.RegularExpressions.Regex.Replace(text, @"\s+", " ");
            return text.Length > 85 ? text.Substring(0, 85) + "..." : text;
        }
        set
        {
            _snippet = value;
            OnPropertyChanged();
        }
    }

    public bool HasAttachments => Attachments.Count > 0;

    public string SenderInitials
    {
        get
        {
            var name = !string.IsNullOrWhiteSpace(FromName) ? FromName : FromEmail;
            if (string.IsNullOrWhiteSpace(name)) return "?";
            
            name = System.Text.RegularExpressions.Regex.Replace(name, @"\<.*?\>|\(.*?\)|\[.*?\]", "").Trim();
            if (string.IsNullOrWhiteSpace(name)) return "?";

            var parts = name.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length >= 2)
            {
                return (parts[0][0].ToString() + parts[parts.Length - 1][0].ToString()).ToUpper();
            }
            return name.Substring(0, Math.Min(2, name.Length)).ToUpper();
        }
    }

    public object? RawXstMessage { get; set; }

    public event PropertyChangedEventHandler? PropertyChanged;
    protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}


public class EmailAttachment
{
    public string FileName { get; set; } = "attachment";
    public byte[] Data { get; set; } = Array.Empty<byte>();
    public long Size => Data.LongLength;

    public string FileTypeIcon
    {
        get
        {
            if (string.IsNullOrWhiteSpace(FileName)) return "generic";
            var ext = System.IO.Path.GetExtension(FileName).ToLowerInvariant().TrimStart('.');
            if (ext == "pdf") return "pdf";
            if (ext == "xls" || ext == "xlsx" || ext == "csv") return "excel";
            if (ext == "doc" || ext == "docx" || ext == "rtf") return "word";
            if (ext == "png" || ext == "jpg" || ext == "jpeg" || ext == "gif" || ext == "bmp") return "image";
            if (ext == "zip" || ext == "rar" || ext == "7z" || ext == "tar" || ext == "gz") return "zip";
            return "generic";
        }
    }

    public string SizeDisplay
    {
        get
        {
            double s = Size;
            string[] units = { "B", "KB", "MB", "GB" };
            int i = 0;
            while (s >= 1024 && i < units.Length - 1) { s /= 1024; i++; }
            return $"{s:0.#} {units[i]}";
        }
    }
}
