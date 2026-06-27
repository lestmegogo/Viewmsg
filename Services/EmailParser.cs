using System.IO;
using MsgViewer.Models;
using MimeMessage = MsgReader.Mime.Message;
using OutlookStorage = MsgReader.Outlook.Storage;

namespace MsgViewer.Services;

/// <summary>
/// Đọc file .msg (Outlook) và .eml (MIME) thành <see cref="EmailMessage"/>.
/// </summary>
public static class EmailParser
{
    public static EmailMessage Parse(string filePath)
    {
        var ext = Path.GetExtension(filePath).ToLowerInvariant();
        return ext switch
        {
            ".msg" => ParseMsg(filePath),
            ".eml" => ParseEml(filePath),
            _ => throw new NotSupportedException($"Định dạng không hỗ trợ: {ext}")
        };
    }

    // ----- .msg (Outlook compound file) -----
    private static EmailMessage ParseMsg(string filePath)
    {
        using var msg = new OutlookStorage.Message(filePath);
        var email = new EmailMessage
        {
            FilePath = filePath,
            Subject = Fallback(msg.Subject),
            Date = msg.SentOn?.DateTime ?? msg.ReceivedOn?.DateTime,
            BodyHtml = msg.BodyHtml,
            BodyText = msg.BodyText,
        };

        if (msg.Sender != null)
        {
            email.FromName = msg.Sender.DisplayName ?? "";
            email.FromEmail = msg.Sender.Email ?? "";
        }

        var to = new List<string>();
        var cc = new List<string>();
        foreach (var r in msg.Recipients)
        {
            var who = FormatRecipient(r.DisplayName, r.Email);
            if (r.Type == MsgReader.Outlook.RecipientType.Cc)
                cc.Add(who);
            else
                to.Add(who);
        }
        email.To = string.Join("; ", to);
        email.Cc = string.Join("; ", cc);

        foreach (var att in msg.Attachments)
        {
            if (att is OutlookStorage.Attachment a)
            {
                email.Attachments.Add(new EmailAttachment
                {
                    FileName = Fallback(a.FileName, "attachment"),
                    Data = a.Data ?? Array.Empty<byte>()
                });
            }
            else if (att is OutlookStorage.Message nested)
            {
                // Email đính kèm bên trong email — lưu dạng .msg
                var name = Fallback(nested.Subject, "embedded") + ".msg";
                email.Attachments.Add(new EmailAttachment { FileName = name });
            }
        }

        return email;
    }

    // ----- .eml (MIME) -----
    private static EmailMessage ParseEml(string filePath)
    {
        var eml = MimeMessage.Load(new FileInfo(filePath));
        var headers = eml.Headers;

        var email = new EmailMessage
        {
            FilePath = filePath,
            Subject = Fallback(headers.Subject),
            Date = headers.DateSent == default ? null : headers.DateSent.DateTime,
        };

        if (headers.From != null)
        {
            email.FromName = headers.From.DisplayName ?? "";
            email.FromEmail = headers.From.Address ?? "";
        }

        email.To = string.Join("; ", headers.To.Select(a => FormatRecipient(a.DisplayName, a.Address)));
        email.Cc = string.Join("; ", headers.Cc.Select(a => FormatRecipient(a.DisplayName, a.Address)));

        email.BodyHtml = eml.HtmlBody?.GetBodyAsText();
        email.BodyText = eml.TextBody?.GetBodyAsText();

        foreach (var att in eml.Attachments)
        {
            email.Attachments.Add(new EmailAttachment
            {
                FileName = Fallback(att.FileName, "attachment"),
                Data = att.Body ?? Array.Empty<byte>()
            });
        }

        return email;
    }

    private static string FormatRecipient(string? name, string? mail)
    {
        if (string.IsNullOrWhiteSpace(name)) return mail ?? "";
        if (string.IsNullOrWhiteSpace(mail)) return name;
        return $"{name} <{mail}>";
    }

    private static string Fallback(string? value, string def = "(Không có tiêu đề)")
        => string.IsNullOrWhiteSpace(value) ? def : value;
}
