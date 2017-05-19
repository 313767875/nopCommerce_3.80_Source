using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Nop.Core.Domain.Messages;
using Nop.Services.Media;
using MimeKit;
using MailKit.Net.Smtp;

namespace Nop.Services.Messages
{
    /// <summary>
    /// Email sender
    /// </summary>
    public partial class EmailSender : IEmailSender
    {
        private readonly IDownloadService _downloadService;

        public EmailSender(IDownloadService downloadService)
        {
            this._downloadService = downloadService;
        }

        /// <summary>
        /// Sends an email
        /// </summary>
        /// <param name="emailAccount">Email account to use</param>
        /// <param name="subject">Subject</param>
        /// <param name="body">Body</param>
        /// <param name="fromAddress">From address</param>
        /// <param name="fromName">From display name</param>
        /// <param name="toAddress">To address</param>
        /// <param name="toName">To display name</param>
        /// <param name="replyTo">ReplyTo address</param>
        /// <param name="replyToName">ReplyTo display name</param>
        /// <param name="bcc">BCC addresses list</param>
        /// <param name="cc">CC addresses list</param>
        /// <param name="attachmentFilePath">Attachment file path</param>
        /// <param name="attachmentFileName">Attachment file name. If specified, then this file name will be sent to a recipient. Otherwise, "AttachmentFilePath" name will be used.</param>
        /// <param name="attachedDownloadId">Attachment download ID (another attachedment)</param>
        /// 
        /// <param name="headers">Headers</param>
        
        public virtual void SendEmail(EmailAccount emailAccount, string subject, string body,
            string fromAddress, string fromName, string toAddress, string toName,
             string replyTo = null, string replyToName = null,
            IEnumerable<string> bcc = null, IEnumerable<string> cc = null,
            string attachmentFilePath = null, string attachmentFileName = null,
            int attachedDownloadId = 0 ,IDictionary < string, string> headers = null)
        {
            var message = new MimeMessage();
            //from, to, reply to
            message.From.Add(new MailboxAddress(fromName, fromAddress));
            message.To.Add(new MailboxAddress(toName, toAddress));
            if (!string.IsNullOrEmpty(replyTo))
            {
                message.ReplyTo.Add(new MailboxAddress(replyToName, replyTo));
            }

            //密件抄送
            if (bcc != null)
            {
                foreach (var address in bcc.Where(bccValue => !string.IsNullOrWhiteSpace(bccValue)))
                {
                    message.Bcc.Add(new MailboxAddress("", address.Trim()));
                }
            }

            //抄送
            if (cc != null)
            {
                foreach (var address in cc.Where(ccValue => !String.IsNullOrWhiteSpace(ccValue)))
                {
                    message.Cc.Add(new MailboxAddress("", address.Trim()));
                }
            }

            var html = new TextPart("html")
            {
                Text = body
            };

            var alternative = new Multipart("alternative");
            alternative.Add(html);
            var multipart = new Multipart("mixed");
            multipart.Add(alternative);

            //添加邮件的附件
            if (!string.IsNullOrEmpty(attachmentFilePath) &&
                File.Exists(attachmentFilePath))
            {

                //附件
                var attachment = new MimePart()
                {
                    ContentObject = new ContentObject(File.OpenRead(attachmentFilePath), ContentEncoding.Default),
                    ContentDisposition = new ContentDisposition(ContentDisposition.Attachment),
                    ContentTransferEncoding = ContentEncoding.Default
                };

                if (!string.IsNullOrEmpty(attachmentFileName))
                {
                    //attachment.FileName = attachmentFileName;
                    attachment.ContentDisposition.Parameters.Add("GB2312", "FileName", attachmentFileName);
                }

                #region  解决附件中文名称乱码情况
                //var index = attachment.Headers.IndexOf(HeaderId.ContentDisposition);
                //var disposition = attachment.Headers[index];
                //disposition.SetValue("GB2312", disposition.Value);
                #endregion

                multipart.Add(attachment);
            }

            //another attachment?
            if (attachedDownloadId > 0)
            {
                var download = _downloadService.GetDownloadById(attachedDownloadId);
                if (download != null)
                {
                    //we do not support URLs as attachments
                    if (!download.UseDownloadUrl)
                    {
                        string fileName = !String.IsNullOrWhiteSpace(download.Filename) ? download.Filename : download.Id.ToString();
                        fileName += download.Extension;

                        var attachment = new MimePart()
                        {
                            ContentObject = new ContentObject(new MemoryStream(download.DownloadBinary), ContentEncoding.Default),
                            ContentDisposition = new ContentDisposition(ContentDisposition.Attachment),
                            ContentTransferEncoding = ContentEncoding.Default
                        };

                        #region  解决附件中文名称乱码情况
                        //var index = attachment.Headers.IndexOf(HeaderId.ContentDisposition);
                        //var disposition = attachment.Headers[index];
                        //disposition.SetValue("GB2312", disposition.Value);
                        attachment.ContentDisposition.Parameters.Add("GB2312", "FileName", fileName);
                        #endregion

                        multipart.Add(attachment);
                    }
                }
            }
            //content
            message.Subject = subject;
            message.Body = multipart;

            //send email

            using (var client = new SmtpClient())
            {
                client.Connect(emailAccount.Host, emailAccount.Port, emailAccount.EnableSsl);
                client.AuthenticationMechanisms.Remove("XOAUTH2");
                client.Authenticate(emailAccount.Username, emailAccount.Password);
                client.Send(message);
                client.Disconnect(true);
            }
        }
    }
}
