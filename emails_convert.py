import win32com.client


def html2oft(
    html_file_path: str,
    otf_file_path: str,
    recipients_to: str = "",
    recipients_cc: str = "",
    message_subject: str = "",
):
    olMailItem = 0x0
    obj = win32com.client.Dispatch("Outlook.Application")

    msg = obj.CreateItem(olMailItem)
    msg.Subject = message_subject
    msg.To = recipients_to
    msg.Cc = recipients_cc
    # olFormatHTML https://msdn.microsoft.com/en-us/library/office/aa219371(v=office.11).aspx
    msg.BodyFormat = 2
    msg.HTMLBody = open(html_file_path).read()
    # newMail.display()
    save_format = 2  # olTemplate	2	Microsoft Outlook template (.oft)
    msg.SaveAs(otf_file_path, save_format)


if __name__ == "__main__":
    import os
    from datetime import datetime

    recipients_to = os.getenv("STREAMING_SDK_EMAIL_RECIPIENTS_TO", "")
    recipients_cc = os.getenv("STREAMING_SDK_EMAIL_RECIPIENTS_CC", "")
    report_date = datetime.today()

    dir = os.getcwd()

    # convert second letter
    first_html_letter = os.path.join(dir, "Letter_1.html")
    if os.path.exists(first_html_letter):
        oft_file = os.path.join(dir, "Letter_1.oft")
        html2oft(
            first_html_letter,
            oft_file,
            message_subject="Streaming SDK Report",
            recipients_to=recipients_to,
            recipients_cc=recipients_cc,
        )
    else:
        print(f"Letter_1.html not found in '{dir}' dir")

    # convert second letter
    second_html_letter = os.path.join(dir, "Letter_2.html")
    if os.path.exists(second_html_letter):
        oft_file = os.path.join(dir, "Letter_2.oft")
        html2oft(
            second_html_letter,
            oft_file,
            message_subject="Weekly QA Report " + report_date.strftime("%d-%b-%Y"),
            recipients_to=recipients_to,
            recipients_cc=recipients_cc,
        )
    else:
        print(f"Letter_2.html not found in '{dir}' dir")
