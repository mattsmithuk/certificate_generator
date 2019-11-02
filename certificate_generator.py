#!/usr/bin/env python3

import os.path, csv, sys
from jinja2 import Environment, FileSystemLoader
from weasyprint import HTML
import yagmail
#import keyring
from gooey import Gooey, GooeyParser


def generate_cert(data, name_input, filename, export_path):
    """
    Generates a PDF certificate using the Full Name and Filename as variables fields
    Date, Location and Organisers have been specified previously in the script
    """

    env = Environment(loader=FileSystemLoader("."))
    template = env.get_template("./template.html")
    try:
        # Setup variables for template
        template_vars = {
            "name": name_input,
            "date": data["date"],
            "location": data["location"],
            "organiser1": data["organiser1"],
            "organiser2": data["organiser2"],
        }
        # Generate HTML
        html_out = template.render(template_vars)
        # Generate PDF
        HTML(string=html_out).write_pdf(
            export_path + filename + ".pdf", stylesheets=["./template.css"]
        )

    except:
        print(f"ERROR: Unable to create PDF {filename}.pdf")


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
            f"\nERROR: Unable to send email to {to_email}. Are you sure you entered the correct password?"
        )
        pass
    else:
        print(f"Certificate sent to {to_email}\n")


@Gooey(
    program_name='ICM Tutorial Certificate Generator',
    program_description='Designed For Use on Windows. GTK Runtime MUST be installed (see File menu)',
    default_size=(610, 750),
    menu=[{
        'name': 'File',
        'items': [{
                'type': 'AboutDialog',
                'menuTitle': 'About',
                'name': 'ICM Tutorial Certificate Generator',
                'description': 'A program to automatically generate and email certificates',
                'version': '0.4',
                'copyright': '2019',
                'website': 'https://github.com/mattsmithuk/certificate_generator/',
                'developer': 'code@mattsmith.email',
                'license': 'MIT'
            }, {
                'type': 'Link',
                'menuTitle': 'GTK Runtime Download',
                'url': 'https://github.com/tschoonj/GTK-for-Windows-Runtime-Environment-Installer/releases/download/2019-07-02/gtk3-runtime-3.24.9-2019-07-02-ts-win64.exe'
            }]
    }]
)
def main():
    """
    Main Function to Generate Certificates (exposed to Gooey)
    """

# Default Email Content
    content = '''Thank you for providing your feedback for this tutorial.\n\nPlease find attached an attendance certificate.'''

    parser = GooeyParser()
    cert_group = parser.add_argument_group('Required')
    email_group = parser.add_argument_group('Optional')

    cert_group.add_argument(
        'filename',
        metavar='Filename',
        help='CSV file with names and email addresses',
        widget='FileChooser',
        gooey_options={
            'validator': {
                'test': 'user_input[-3:] == "csv"',
                'message': 'Filetype must be CSV'
            }
        })

    cert_group.add_argument('export_path', metavar='Export Path', help='Directory to save certificates to', widget='DirChooser')
    cert_group.add_argument('date', metavar='Tutorial Date', widget='TextField')
    cert_group.add_argument('location', metavar='Tutorial Location', widget='TextField')
    cert_group.add_argument('organiser1', metavar='First Organiser', widget='TextField')
    cert_group.add_argument('--organiser2', metavar='Second Organiser (optional)', widget='TextField')
    
    email_group.add_argument('--send', metavar='Send Emails', help='Select to Send Emails Automatically', action='store_true')
    email_group.add_argument('--email', metavar='Email Address to Send From', help="Must be a Gmail Account")
    email_group.add_argument('--password', metavar='Password', widget='PasswordField')

    args = parser.parse_args()

    working_dir = args.export_path + "\\"

    if args.send:
        yagmail.register(args.email, args.password)

    data = {}
    data["date"] = args.date
    data["location"] = args.location
    data["organiser1"] = args.organiser1
    data["organiser2"] = args.organiser2
    data["email"] = args.email
    data["password"] = args.password
    
    i = 0
  
    with open(args.filename, mode="r") as f:
        f_csv = csv.DictReader(f)
        for csv_row in f_csv:
            i += 1
            pdf_filename = f"{i}_{csv_row['First Name'][0]+csv_row['Last Name'][0]}_Certificate"
            print(f"DEBUG: Filename is {pdf_filename}.pdf")
            full_name = csv_row["First Name"] + " " + csv_row["Last Name"]
            print(f"DEBUG: Full name is {full_name}")

            generate_cert(data, full_name, pdf_filename, working_dir)
    
            if args.send:
                print(f"Sending Certificate to {csv_row['Email']}")
                send_email(
                    email_user=data["email"],
                    email_pass=data["password"],
                    to_email=csv_row["Email"],
                    firstname=csv_row["First Name"],
                    date=data["date"],
                    msg_content=content,
                    attachment=working_dir + pdf_filename + ".pdf",
            )

if __name__ == "__main__":
    main()