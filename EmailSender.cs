using ClosedXML.Excel;
using MailKit.Net.Smtp;
using MimeKit;

namespace EmailSenderApp
{
    internal class EmailSender
    {
        private readonly string bodyPath = @"..\..\..\Email-body\email-body.txt";

        private readonly string successPath = @"..\..\..\Email-status\Success.xlsx";
        private readonly string failedPath = @"..\..\..\Email-status\Failed.xlsx";

        private readonly Dictionary<string, (string techKeyword, string resumePath, string emailFile)> sources = new()
        {
            { "general", ("Blazor, Angular, and React", @"..\..\..\Resume\Mahadev_Pimpalkar_Dotnet_Developer.pdf", @"..\..\..\Emails-list\general-recruiters-email-list.txt") },
            { "blazor",  ("Blazor", @"..\..\..\Resume\Mahadev_Pimpalkar_Dotnet_Developer_Blazor.pdf", @"..\..\..\Emails-list\Blazor\blazor-recruiters-email-list.txt") },
            { "angular", ("Angular", @"..\..\..\Resume\Mahadev_Pimpalkar_Dotnet_Developer_Angular.pdf", @"..\..\..\Emails-list\Angular\angular-recruiters-email-list.txt") },
            { "react",   ("ReactJS", @"..\..\..\Resume\Mahadev_Pimpalkar_Dotnet_Developer_React.pdf", @"..\..\..\Emails-list\React\react-recruiters-email-list.txt") }
        };

        public async Task RunAsync()
        {
            CreateOrStyleExcel(successPath);
            CreateOrStyleExcel(failedPath);

            foreach (var source in sources.Values)
            {
                var emailFile = source.emailFile;
                var tech = source.techKeyword;
                var resume = source.resumePath;
                var processedEmails = new List<string>();

                if (!File.Exists(emailFile)) continue;

                var emails = File.ReadAllLines(emailFile);
                if (emails.Length == 0) continue;

                var rawBody = File.ReadAllText(bodyPath);
                var emailBody = rawBody.Replace("{TECHSTACK}", tech);
                
                string subject = string.Empty;
                if (tech == "Blazor, Angular, and React")
                {
                     subject = "Application for the .Net Developer - Immediate Joiner";

                }
                else
                {
                     subject = $"Application for the .Net {tech} Developer - Immediate Joiner";
                }


                foreach (var email in emails)
                {
                    MimeMessage message;

                    try
                    {
                        message = CreateEmailMessage("Mahadev Pimpalkar", "mahadeopimpalkar16@gmail.com", email, subject, emailBody, resume);
                    }
                    catch (FormatException)
                    {
                        Console.WriteLine($"Invalid email format: {email}");
                        LogToExcel(tech, failedPath, email, "Invalid email format");
                        processedEmails.Add(email);
                        continue;
                    }

                    await SendEmailAsync(tech, message, email, resume, processedEmails);
                    await Task.Delay(1000);
                }

                RemoveProcessedEmails(emailFile, processedEmails);
            }
        }

        private void RemoveProcessedEmails(string path, List<string> sentEmails)
        {
            var all = File.ReadAllLines(path).ToList();
            var pending = all.Except(sentEmails).ToList();
            File.WriteAllLines(path, pending);
            Console.WriteLine($"Processed emails removed from {Path.GetFileName(path)}");
        }

        private MimeMessage CreateEmailMessage(string senderName, string senderEmail, string recipientEmail, string subject, string body, string attachmentPath)
        {
            var message = new MimeMessage();
            message.From.Add(new MailboxAddress(senderName, senderEmail));
            message.Subject = subject;
            message.To.Add(MailboxAddress.Parse(recipientEmail));

            var builder = new BodyBuilder { TextBody = body };
            builder.Attachments.Add(attachmentPath);
            message.Body = builder.ToMessageBody();

            return message;
        }

        private async Task SendEmailAsync( string TechStack, MimeMessage message, string email, string attachmentPath, List<string> processedList)
        {
            using var client = new SmtpClient();

            try
            {
                await client.ConnectAsync("smtp.gmail.com", 587, false);
                await client.AuthenticateAsync("mahadeopimpalkar16@gmail.com", "ldipekvdwcvdhxhe");
                await client.SendAsync(message);
                Console.WriteLine($"{email} : Sent successfully");
                LogToExcel(TechStack, successPath, email, "Sent successfully");
            }
            catch (SmtpCommandException ex) when (ex.ErrorCode == SmtpErrorCode.RecipientNotAccepted)
            {
                Console.WriteLine($"{email} : Address not found");
                LogToExcel(TechStack, failedPath, email, "Address not found");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"{email} : General failure - {ex.Message}");
                LogToExcel(TechStack, failedPath, email, $"General failure: {ex.Message}");
            }
            finally
            {
                await client.DisconnectAsync(true);
                processedList.Add(email);
            }
        }

        private void LogToExcel(string tecchStack, string filePath, string email, string statusMessage)
        {
            if(tecchStack.Contains("Blazor, Angular, and React"))
            {
                tecchStack = "General";
            }

            var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

            XLWorkbook workbook = File.Exists(filePath) ? new XLWorkbook(filePath) : new XLWorkbook();
            IXLWorksheet sheet = workbook.Worksheets.FirstOrDefault() ?? workbook.AddWorksheet("Sheet1");

            int row = sheet.LastRowUsed()?.RowNumber() ?? 1;
            row += 1;

            sheet.Cell(row,1).Value = tecchStack;
            sheet.Cell(row, 2).Value = email;
            sheet.Cell(row, 3).Value = statusMessage;
            sheet.Cell(row, 4).Value = timestamp;

            var Day = DateTime.Now.Day;
            var rowStyle = sheet.Range(row, 1, row, 4);

            var setColor = PopulateColor();
            if (statusMessage.ToLower().Contains("success"))
                rowStyle.Style.Fill.BackgroundColor = setColor;
            else if (statusMessage.ToLower().Contains("invalid") || statusMessage.ToLower().Contains("address not found"))
                rowStyle.Style.Fill.BackgroundColor = Day % 2 == 0 ? XLColor.MistyRose : XLColor.LightPink;
            else
                rowStyle.Style.Fill.BackgroundColor = XLColor.LightGoldenrodYellow;

            rowStyle.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            rowStyle.Style.Border.InsideBorder = XLBorderStyleValues.Thin;


            var header = sheet.Range("A1:D1");
            header.Style.Font.Bold = true;
            header.Style.Fill.BackgroundColor = XLColor.LightBlue;
            header.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            header.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            // Border
            header.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            header.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            workbook.SaveAs(filePath);
        }

        private void CreateOrStyleExcel(string filePath)
        {
            if (!File.Exists(filePath))
            {
                var workbook = new XLWorkbook();
                var sheet = workbook.AddWorksheet("Sheet1");
                sheet.Cell("A1").Value = "TechStack";
                sheet.Cell("B1").Value = "Email";
                sheet.Cell("C1").Value = "Status";
                sheet.Cell("D1").Value = "Date";

                var header = sheet.Range("A1:D1");
                header.Style.Font.Bold = true;
                header.Style.Fill.BackgroundColor = XLColor.LightBlue;

                workbook.SaveAs(filePath);
            }
        }

        public XLColor PopulateColor()
        {
            DayOfWeek day = new DayOfWeek();
            return day switch
            {
                DayOfWeek.Sunday => XLColor.LightApricot,
                DayOfWeek.Monday => XLColor.LightGreen,
                DayOfWeek.Tuesday => XLColor.LightGray,
                DayOfWeek.Wednesday => XLColor.LightMauve,
                DayOfWeek.Thursday => XLColor.LightBlue,
                DayOfWeek.Friday => XLColor.LightCyan,
                DayOfWeek.Saturday => XLColor.LightCoral,
                 _ => XLColor.Linen
            };
        }
    }
}
