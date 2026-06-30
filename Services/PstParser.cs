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
    public static EmailMessage MapMessageSummary(XstMessage msg, string pstFilePath)
    {
        var email = new EmailMessage
        {
            FilePath = pstFilePath + "||" + (msg.Path ?? Guid.NewGuid().ToString()),
            Subject = string.IsNullOrWhiteSpace(msg.Subject) ? "(Không có tiêu đề)" : msg.Subject,
            Date = msg.Date ?? msg.ReceivedTime ?? msg.SubmittedTime,
            IsRead = msg.IsRead,
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
                // Thử đọc HTML thô từ thuộc tính MAPI đối với RTF có nhúng HTML
                var htmlProp = msg.Properties[PropertyCanonicalName.PidTagBodyHtml];
                if (htmlProp?.Value != null)
                {
                    if (htmlProp.Value is byte[] htmlBytes)
                        bodyHtml = msg.Encoding.GetString(htmlBytes);
                    else
                        bodyHtml = htmlProp.Value.ToString();
                }
                else
                {
                    bodyText = body.Text;
                }
            }
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
