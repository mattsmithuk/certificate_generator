#!/usr/bin/env python3

import csv
import os
import sys
from gooey import Gooey, GooeyParser


def generate_cert(args, name_input, filename, export_path):
    import subprocess
    import jinja2

    template_path = resource_path("incl\\template.html")
    context = {
        "name": name_input,
        "date": args.date,
        "location": args.location,
        "organiser1": args.organiser1,
        "organiser2": args.organiser1,
    }

    def render(template_path, context):
        # Context = jinja context dictionary
        path, filename = os.path.split(template_path)
        return (
            jinja2.Environment(loader=jinja2.FileSystemLoader(path or "./"))
            .get_template(filename)
            .render(context)
        )

    def write_html(html_name, context):
        with open(html_name, "w") as f:
            html = render(template_path, context)
            f.write(html)

    def create_pdf(args, template_path, out_dir, context, keep_html=False):
        html_name = os.path.join(export_path, filename + ".html")
        pdf_name = os.path.join(export_path, filename + ".pdf")
        wkhtmltopdf_path = resource_path("incl\\wkhtmltopdf.exe")

        write_html(html_name, context)
        if args.verbose:
            print(f"DEBUG: Created {html_name}")

        sub_args = f'"{wkhtmltopdf_path}" --zoom 1.244 -B 0 -L 1 -R 0 -T 10 "{html_name}" "{pdf_name}"'
        child = subprocess.Popen(
            sub_args, shell=True, stderr=subprocess.DEVNULL, stdout=subprocess.DEVNULL
        )
        # Wait for the sub-process to terminate (communicate) then find out the
        # status code
        output, errors = child.communicate()
        if child.returncode or errors:
            # Change this to a log message
            print(
                f"ERROR: PDF conversion failed. Subprocess exit status {child.returncode}"
            )
        else:
            if args.verbose:
                print(f"DEBUG: Created {pdf_name}")

        if not keep_html:
            # Try block only to avoid raising an error if you've already moved
            # or deleted the html
            try:
                os.remove(html_name)
            except OSError:
                pass

        return pdf_name

    return create_pdf(args, template_path, export_path, context)


def send_gmail(sender, password, recipient, subject, message, attach_file=None):
    import sys
    import smtplib
    import time
    import imaplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.base import MIMEBase
    from email import encoders

    msg = MIMEMultipart()
    msg["From"] = sender
    msg["To"] = recipient
    msg["Subject"] = subject

    afilepath, afilename = os.path.split(attach_file)

    msg.attach(MIMEText(message, "plain"))

    # attachment
    if attach_file != None:
        filename = afilename
        attachment = open(attach_file, "rb")

        part = MIMEBase("application", "octet-stream")
        part.set_payload((attachment).read())
        encoders.encode_base64(part)
        part.add_header(
            f"Content-Disposition", "attachment; filename= {}".format(afilename)
        )

        msg.attach(part)

    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login(sender, password)
    text = msg.as_string()
    server.sendmail(sender, recipient, text)
    server.quit()


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


@Gooey(
    program_name="ICM Tutorial Certificate Generator",
    program_description="Designed For Use on Windows",
    default_size=(680, 750),
    menu=[
        {
            "name": "File",
            "items": [
                {
                    "type": "AboutDialog",
                    "menuTitle": "About",
                    "name": "ICM Tutorial Certificate Generator",
                    "description": "A program to automatically generate and email certificates",
                    "version": "0.5",
                    "copyright": "2019",
                    "website": "https://github.com/mattsmithuk/certificate_generator/",
                    "developer": "code@mattsmith.email",
                    "license": "MIT",
                },
            ],
        },
        {
            "name": "Help",
            "items": [
                {
                    "type": "MessageDialog",
                    "menuTitle": "CSV File",
                    "message": 'The CSV file MUST include the following columns: "First Name", "Last Name" and "Email"',
                    "caption": "CSV File Requirements",
                }
            ],
        },
    ],
)
def main():
    """
    Main Function to Generate Certificates (exposed to Gooey)
    """

    # Default Email Content
    content = """Thank you for providing your feedback for this tutorial.\n\nPlease find attached an attendance certificate."""

    # Collect Input Arguements
    parser = GooeyParser()
    cert_group = parser.add_argument_group("Required")
    email_group = parser.add_argument_group("Optional")

    cert_group.add_argument(
        "filename",
        metavar="Filename",
        help="CSV file with names and email addresses",
        widget="FileChooser",
        gooey_options={
            "validator": {
                "test": 'user_input[-3:] == "csv"',
                "message": "Filetype must be CSV",
            }
        },
    )

    cert_group.add_argument(
        "export_path",
        metavar="Export Path",
        help="Directory to save certificates to",
        widget="DirChooser",
    )
    cert_group.add_argument("date", metavar="Tutorial Date", widget="TextField")
    cert_group.add_argument("location", metavar="Tutorial Location", widget="TextField")
    cert_group.add_argument("organiser1", metavar="First Organiser", widget="TextField")
    cert_group.add_argument(
        "--organiser2", metavar="Second Organiser (optional)", widget="TextField"
    )
    email_group.add_argument(
        "--verbose",
        metavar="Verbose Output",
        help="Select to see detailed output",
        action="store_true",
    )
    email_group.add_argument(
        "--send",
        metavar="Send Emails",
        help="Select to Send Emails Automatically",
        action="store_true",
    )

    email_group.add_argument(
        "--email", metavar="Email Address to Send From", help="Must be a Gmail Account"
    )
    email_group.add_argument("--password", metavar="Password", widget="PasswordField")

    # Process inputs
    args = parser.parse_args()

    # Iterate through CSV file
    i = 0
    with open(args.filename, mode="r") as f:
        f_csv = csv.DictReader(f)
        for csv_row in f_csv:
            i += 1
            output_filename = (
                f"{i}_{csv_row['First Name'][0]+csv_row['Last Name'][0]}_Certificate"
            )

            full_name = csv_row["First Name"] + " " + csv_row["Last Name"]

            if args.verbose:
                print(f"DEBUG: Filename will be {output_filename}.pdf")
                print(f"DEBUG: Full name is {full_name}")

            pdf_filename = generate_cert(
                args, full_name, output_filename, args.export_path
            )

            # Send Emails if selected
            if args.send:
                print(f"MSG: Sending Certificate to {csv_row['Email']}")
                try:
                    send_gmail(
                        sender=args.email,
                        password=args.password,
                        recipient=csv_row["Email"],
                        subject=f"Attendance Certificate for Tutorial on {args.date}",
                        message=content,
                        attach_file=pdf_filename,
                    )
                except:
                    print(
                        f'ERROR: Unable to send email to {csv_row["Email"]}. Are you sure you entered the correct password?'
                    )
                    pass
                else:
                    print(f'MSG: Certificate sent to {csv_row["Email"]}\n')


if __name__ == "__main__":
    main()
