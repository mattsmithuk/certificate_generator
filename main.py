#!/usr/bin/env python3

import csv
from gooey import Gooey, GooeyParser
import yagmail


def generate_cert(args, name_input, filename, export_path):
    import subprocess
    import jinja2
    import os

    template_path = "./template.html"
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

        write_html(html_name, context)
        if args.verbose:
            print(f"DEBUG: Created {html_name}")

        sub_args = f'".\\wkhtmltopdf.exe" --zoom 1.244 -B 0 -L 1 -R 0 -T 10 "{html_name}" "{pdf_name}"'
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


def send_email(
    email_user, email_pass, to_email, firstname, date, msg_content, attachment
):
    """
    Sends an email from a gmail account
    Uses 'msg_template' as an email template using string formatting
    """

    yag = yagmail.SMTP(email_user, email_pass)

    try:
        yag.send(
            to=to_email,
            subject=f"Attendance Certificate for Tutorial on {date}",
            contents=msg_content.format(firstname),
            attachments=attachment,
        )
    except:
        print(
            f"ERROR: Unable to send email to {to_email}. Are you sure you entered the correct password?"
        )
        pass
    else:
        print(f"MSG: Certificate sent to {to_email}\n")


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
        }
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

    # Setup Yagmail if 'Send Emails' selected
    if args.send:
        yagmail.register(args.email, args.password)

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
                send_email(
                    email_user=args.email,
                    email_pass=args.password,
                    to_email=csv_row["Email"],
                    firstname=csv_row["First Name"],
                    date=args.date,
                    msg_content=content,
                    attachment=pdf_filename,
                )


if __name__ == "__main__":
    main()
