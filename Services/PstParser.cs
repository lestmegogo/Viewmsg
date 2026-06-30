using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using XstReader;
using MsgViewer.Models;
using XstReader.ElementProperties;

namespace MsgViewer.Services;

/// <summary>
/// Parser chuyên biệt để đọc và ánh xạ dữ liệu từ file .pst/.ost (XstReader) sang mô hình EmailMessage chung.
/// </summary>
public static class PstParser
{
    public static string GetMessageKey(XstMessage msg)
    {
        if (!string.IsNullOrWhiteSpace(msg.Path))
            return msg.Path;

        string datePart = msg.Date?.Ticks.ToString() ?? msg.ReceivedTime?.Ticks.ToString() ?? "nodate";
        string subjectPart = msg.Subject ?? "";
        string fromPart = msg.From ?? "";
        string rawKey = $"{subjectPart}_{fromPart}_{datePart}";
        using (var sha1 = System.Security.Cryptography.SHA1.Create())
        {
            byte[] hash = sha1.ComputeHash(System.Text.Encoding.UTF8.GetBytes(rawKey));
            return "hash_" + BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
        }
    }

    public static EmailMessage MapMessageSummary(XstMessage msg, string pstFilePath)
    {
        string? snippet = null;
        try
        {
            var bodyProp = msg.Properties[PropertyCanonicalName.PidTagBody]?.Value?.ToString();
            if (!string.IsNullOrWhiteSpace(bodyProp))
            {
                snippet = bodyProp;
            }
            else
            {
                var bodyObj = msg.Body ?? msg.GetBody();
                if (bodyObj != null && !string.IsNullOrWhiteSpace(bodyObj.Text))
                {
                    snippet = bodyObj.Text;
                }
            }
        }
        catch { }

        var email = new EmailMessage
        {
            FilePath = pstFilePath + "||" + GetMessageKey(msg),
            Subject = string.IsNullOrWhiteSpace(msg.Subject) ? "(Không có tiêu đề)" : msg.Subject,
            Date = msg.Date ?? msg.ReceivedTime ?? msg.SubmittedTime,
            IsRead = msg.IsRead,
            BodyText = snippet
        };

        email.FromName = msg.From ?? "";
        
        // Trích xuất Email người gửi từ thuộc tính MAPI (vì XstMessage.From thường chỉ chứa Tên hiển thị)
        var senderEmail = msg.Properties[PropertyCanonicalName.PidTagSenderEmailAddress]?.Value?.ToString();
        if (string.IsNullOrWhiteSpace(senderEmail))
        {
            // Fallback: Tìm trong danh sách người gửi của Recipients
            var senderRecip = msg.Recipients[RecipientType.Sender].FirstOrDefault() 
                              ?? msg.Recipients[RecipientType.SentRepresenting].FirstOrDefault();
            senderEmail = senderRecip?.Address;
        }
        email.FromEmail = senderEmail ?? "";

        // Người nhận To và Cc hiển thị
        email.To = msg.To ?? "";
        email.Cc = msg.Cc ?? "";

        // Ánh xạ danh sách tệp đính kèm sơ bộ (chỉ tên tệp để hiển thị chip đính kèm ở danh sách thư)
        if (msg.Attachments != null)
        {
            foreach (var att in msg.Attachments)
            {
                if (att != null && !string.IsNullOrWhiteSpace(att.FileName))
                {
                    if (att.IsInlineAttachment || att.IsHidden) continue;
                    email.Attachments.Add(new EmailAttachment { FileName = att.FileName });
                }
            }
        }

        // Lưu giữ liên kết đến đối tượng tin nhắn gốc phục vụ cho lazy load nội dung chi tiết
        email.RawXstMessage = msg;

        return email;
    }

    public static void LoadMessageDetails(EmailMessage email)
    {
        if (email.RawXstMessage is not XstMessage msg) return;

        // Clear đính kèm cũ trước khi nạp lại
        email.Attachments.Clear();

        // 1. Nạp nội dung thư (Lazy Loading Body)
        string? bodyHtml = null;
        string? bodyText = null;

        try
        {
            var body = msg.Body ?? msg.GetBody();
            if (body != null)
            {
                if (body.Format == XstMessageBodyFormat.Html)
                {
                    bodyHtml = body.Text;
                }
                else if (body.Format == XstMessageBodyFormat.PlainText)
                {
                    bodyText = body.Text;
                }
                else if (body.Format == XstMessageBodyFormat.Rtf)
                {
                    var htmlProp = msg.Properties[PropertyCanonicalName.PidTagBodyHtml];
                    if (htmlProp?.Value != null)
                    {
                        if (htmlProp.Value is byte[] htmlBytes)
                        {
                            var encoding = msg.Encoding ?? System.Text.Encoding.UTF8;
                            bodyHtml = encoding.GetString(htmlBytes);
                        }
                        else
                        {
                            bodyHtml = htmlProp.Value.ToString();
                        }
                    }
                    else
                    {
                        bodyText = body.Text;
                    }
                }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Failed to load standard PST body: {ex.Message}");
        }

        // Fallback cực mạnh nếu không đọc được body theo cách thông thường (đặc biệt khi giải nén RTF bị lỗi hoặc Encoding bị null)
        if (string.IsNullOrWhiteSpace(bodyHtml) && string.IsNullOrWhiteSpace(bodyText))
        {
            try
            {
                var htmlProp = msg.Properties[PropertyCanonicalName.PidTagBodyHtml];
                if (htmlProp?.Value != null)
                {
                    if (htmlProp.Value is byte[] htmlBytes)
                    {
                        var encoding = msg.Encoding ?? System.Text.Encoding.UTF8;
                        bodyHtml = encoding.GetString(htmlBytes);
                    }
                    else
                    {
                        bodyHtml = htmlProp.Value.ToString();
                    }
                }

                if (string.IsNullOrWhiteSpace(bodyHtml))
                {
                    var textProp = msg.Properties[PropertyCanonicalName.PidTagBody];
                    if (textProp?.Value != null)
                    {
                        if (textProp.Value is byte[] textBytes)
                        {
                            var encoding = msg.Encoding ?? System.Text.Encoding.UTF8;
                            bodyText = encoding.GetString(textBytes);
                        }
                        else
                        {
                            bodyText = textProp.Value.ToString();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to fallback load PST body properties: {ex.Message}");
            }
        }

        if (string.IsNullOrWhiteSpace(bodyText) && !string.IsNullOrWhiteSpace(bodyHtml))
        {
            var cleanText = System.Text.RegularExpressions.Regex.Replace(bodyHtml, "<.*?>", string.Empty);
            cleanText = System.Net.WebUtility.HtmlDecode(cleanText);
            bodyText = cleanText;
        }

        email.BodyHtml = bodyHtml;
        email.BodyText = bodyText;

        // 2. Nạp danh sách tệp đính kèm (Lazy Loading Attachments)
        var attachments = msg.Attachments ?? msg.GetAttachments();
        if (attachments != null)
        {
            foreach (var att in attachments)
            {
                if (att.IsInlineAttachment || att.IsHidden) continue;
                if (att.IsFile)
                {
                    var emailAtt = new EmailAttachment
                    {
                        FileName = att.DisplayName ?? att.FileNameForSaving ?? att.FileName ?? "attachment"
                    };

                    try
                    {
                        using (var ms = new MemoryStream())
                        {
                            att.SaveToStream(ms);
                            emailAtt.Data = ms.ToArray();
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Failed to load data for attachment {emailAtt.FileName}: {ex.Message}");
                        emailAtt.Data = Array.Empty<byte>();
                    }

                    email.Attachments.Add(emailAtt);
                }
                else if (att.IsEmail && att.AttachedEmailMessage != null)
                {
                    // Email nằm lồng trong email
                    var nestedSubject = att.AttachedEmailMessage.Subject ?? "embedded_message";
                    var emailAtt = new EmailAttachment
                    {
                        FileName = nestedSubject + ".msg",
                        Data = Array.Empty<byte>() // Đánh dấu không có dữ liệu nhị phân trực tiếp
                    };
                    email.Attachments.Add(emailAtt);
                }
            }
        }
    }
}
