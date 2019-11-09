import os.path, csv, sys, json
from jinja2 import Environment, FileSystemLoader
from weasyprint import HTML
import yagmail
import keyring
from tkinter import Tk, messagebox
from tkinter.filedialog import askopenfilename

def generate_cert(data, name_input, filename, export_path):
    """
    Generates a PDF certificate using the Full Name and Filename as variables fields
    Date, Location and Organisers have been specified previously in the script
    """

    env = Environment(loader=FileSystemLoader("."))
    template = env.get_template("./assets/template/template.html")
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
        render_pdf_phantomjs(html_out)

    except:
        print(f"ERROR: Unable to create PDF {filename}.pdf")
        messagebox.showinfo("Error", f"Unable to create PDF {filename}.pdf")

def render_pdf_phantomjs(html):
    """mimerender helper to render a PDF from HTML using phantomjs."""
    # The 'makepdf.js' PhantomJS program takes HTML via stdin and returns PDF binary via stdout
    # https://gist.github.com/philfreo/5854629
    from subprocess import Popen, PIPE, STDOUT
    import os
    p = Popen(['phantomjs', '%s/makepdf.js' % os.path.dirname(os.path.realpath(__file__))], stdout=PIPE, stdin=PIPE, stderr=STDOUT)
    return p.communicate(input=html.encode('utf-8'))[0]

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
        messagebox.showinfo(
            "Error",
            f"Unable to send email to {to_email}. Are you sure you entered the correct password?",
        )
        pass
    else:
        print(f"Certificate sent to {to_email}\n")


def ask_file(file_type):
    """
    Ask the user to select a file
    Taken from auto-py-to-exe
    """
    root = Tk()
    root.withdraw()
    root.wm_attributes("-topmost", 1)

    if file_type == "python":
        file_types = [("Python files", "*.py;*.pyw"), ("All files", "*")]
    elif file_type == "icon":
        file_types = [("Icon files", "*.ico"), ("All files", "*")]
    elif file_type == "json":
        file_types = [("JSON Files", "*.json"), ("All files", "*")]
    elif file_type == "csv":
        file_types = [("Comma-seperated values", "*.csv"), ("All files", "*")]
    else:
        file_types = [("All files", "*")]
    file_path = askopenfilename(parent=root, filetypes=file_types)
    root.update()
    return file_path


def main(json_data):
    """
    Main Function to Generate Certificates
    """
    data = json.loads(json_data, strict=False)
    working_dir = os.path.dirname(data["data_file"]) + "\\"

    print(f"DEBUG: Received data type: {type(json_data)}")
    print("DEBUG: Received data:")
    print(json.dumps(json_data, indent=4, sort_keys=True))

    yagmail.register(data["email"], data["password"])

    i = 0

    try:
        with open(data["data_file"], mode="r") as f:
            f_csv = csv.DictReader(f)

            for csv_row in f_csv:
                i += 1
                filename = f"{i}_{csv_row['First Name'][0]+csv_row['Last Name'][0]}_Certificate"
                print(f"DEBUG: Filename is {filename}")
                full_name = csv_row["First Name"] + " " + csv_row["Last Name"]
                print(f"DEBUG: Full name is {full_name}")

                generate_cert(data, full_name, filename, working_dir)

                print(f"Sending Certificate to {csv_row['Email']}")

                send_email(
                    email_user=data["email"],
                    email_pass=data["password"],
                    to_email=csv_row["Email"],
                    firstname=csv_row["First Name"],
                    date=data["date"],
                    msg_content=data["content"],
                    attachment=working_dir + filename + ".pdf",
                )

    except KeyError:
        print("ERROR: CSV not formatted correctly")
        messagebox.showinfo("CSV File Error", "CSV not formatted correctly\nColumns must be 'First Name', 'Last Name' and 'Email'")
        sys.exit()

def main():
    print("Hello")