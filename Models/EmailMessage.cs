using System.Collections.Generic;

namespace MsgViewer.Models;

/// <summary>
/// Mô hình email thống nhất cho cả file .msg và .eml.
/// </summary>
public class EmailMessage
{
    public string FilePath { get; set; } = "";
    public string Subject { get; set; } = "(Không có tiêu đề)";
    public string FromName { get; set; } = "";
    public string FromEmail { get; set; } = "";
    public string To { get; set; } = "";
    public string Cc { get; set; } = "";
    public DateTime? Date { get; set; }
    public bool IsRead { get; set; } = true;
    public bool IsSelected { get; set; } = false;

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

    /// <summary>Nội dung HTML (ưu tiên). Có thể rỗng.</summary>
    public string? BodyHtml { get; set; }

    /// <summary>Nội dung text thuần (fallback khi không có HTML).</summary>
    public string? BodyText { get; set; }

    public List<EmailAttachment> Attachments { get; } = new();

    /// <summary>Tên người gửi hiển thị: "Tên &lt;email&gt;" hoặc chỉ email.</summary>
    public string FromDisplay =>
        string.IsNullOrWhiteSpace(FromName)
            ? FromEmail
            : string.IsNullOrWhiteSpace(FromEmail) ? FromName : $"{FromName} <{FromEmail}>";

    public string FromClean => string.IsNullOrWhiteSpace(FromName) ? FromEmail : FromName;

    public string DateDisplay => Date?.ToString("dddd, dd/MM/yyyy HH:mm") ?? "";

    public string Snippet
    {
        get
        {
            if (string.IsNullOrWhiteSpace(BodyText)) return "";
            var text = BodyText.Replace("\r", " ").Replace("\n", " ").Trim();
            text = System.Text.RegularExpressions.Regex.Replace(text, @"\s+", " ");
            return text.Length > 85 ? text.Substring(0, 85) + "..." : text;
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

    /// <summary>Đối tượng XstMessage gốc (chỉ dùng cho tệp PST/OST để nạp lazy).</summary>
    public object? RawXstMessage { get; set; }
}


public class EmailAttachment
{
    public string FileName { get; set; } = "attachment";
    public byte[] Data { get; set; } = Array.Empty<byte>();
    public long Size => Data.LongLength;

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
