import sys
import os
import smtplib
import yaml
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from PyQt5.QtGui import QTextCursor, QTextDocument
from PyQt5.QtCore import Qt, QEvent, QObject, QSize, QRect
from PyQt5.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QComboBox,
    QCompleter,
    QTextEdit,
    QPushButton,
    QTabWidget,
    QPlainTextEdit,
    QToolBar,
    QAction,
)
from jinja2 import Environment, FileSystemLoader, meta
import datetime
from plyer import notification
# -------------------------- Function for eval code in mail --------
def e_date():
    return datetime.date.today().strftime('%Y-%m-%d')
def e_signature():
    return """Signature
"""
# ----------------------------------------------------------------
# -------------------------- HTML Base of mail -------------------
toHtml = lambda x:  f"""
<!DOCTYPE html>
<html lang="en">
<head>
    <head><meta name="qrichtext" content="1" />
    <style>
        body{{
            padding: 20px;
        }}
        .notify{{
            background-color: #f2f2f2;
            padding: 10px;
            margin-bottom: 20px;
        }}
        /**
        * Add your custom styles here
        */
    </style>

</head>
<body>
    {x}
</body>
</html>
"""
# ----------------------------------------------------------------
def extractBody(html):
    return html.split("<body>")[1].split("</body>")[0]
class CustomCompleter(QCompleter):
    def __init__(self,wordlist=[], parent=None):
        super(CustomCompleter, self).__init__(wordlist,parent)
        self.activated.connect(self.insert_completion)

    def insert_completion(self, completion):
        if self.widget() is not None:
            cursor = self.widget().textCursor()
            extra = len(completion) - len(self.completionPrefix())
            cursor.movePosition(QTextCursor.Left)
            cursor.movePosition(QTextCursor.EndOfWord)
            cursor.insertText(completion[-extra:])
            self.widget().setTextCursor(cursor)

class EmailApp(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Email App")
        self.setGeometry(100, 100, 800, 600)
        self.var_inputs = {}
        self.layout_ = None 
        self.input_layout = None
        self.templates_dir = "template"
        self.config_file = "config.yaml"
        self.dest_file = "dest.txt"
        self.emails = []
        self.toolbar  = None
        self.current_template = ""
        self.templates = []
        self.templatesData = {}
        self.smtp_config = {}
        self.vars = {}
        self.load_smtp_config()
        self.load_templates()
        self.load_dest()

        self.init_ui()

    def load_smtp_config(self):
        with open(self.config_file, "r") as file:
            self.smtp_config = yaml.safe_load(file)
        for k in self.smtp_config:
            self.emails.append(k + "=>" + self.smtp_config[k]["user"])

    def load_dest(self):
        with open(self.dest_file, "r") as file:
            self.destinations = file.read().split("\n")

    def load_templates(self):
        self.templates = []
        for file_name in os.listdir(self.templates_dir):
            if file_name.endswith(".html"):
                with open(os.path.join(self.templates_dir, file_name), "r") as file_:
                    self.templates.append(file_name)
                    self.templatesData[file_name] = {"content": file_.read()}
                with open(
                    os.path.join(self.templates_dir, file_name.replace(".html", ".yaml")),
                    "r",
                ) as file_:
                    data = yaml.safe_load(file_)
                    self.templatesData[file_name].update(data)

    def init_ui(self):
        main_tab_widget = QTabWidget(self)

        original_tab = QWidget()
        self.setup_original_tab(original_tab)
        main_tab_widget.addTab(original_tab, "Send new Mail")

        new_template_tab = QWidget()
        self.setup_new_template_tab(new_template_tab)
        main_tab_widget.addTab(new_template_tab, "New Template Creation")
        
        edit_confif_tab = QWidget()
        self.setup_edit_confif_tab(edit_confif_tab)
        main_tab_widget.addTab(edit_confif_tab, "Edit Configurations")
        if len(self.emails) == 0:
            main_tab_widget.setTabEnabled(0, False)
            main_tab_widget.currentIndex = 2

        main_tab_widget.currentChanged.connect(self.on_tab_changed)
        self.setCentralWidget(main_tab_widget)

    def setup_original_tab(self, tab_widget):
        layout = QVBoxLayout(tab_widget)

        email_label = QLabel("Select Email:")
        self.email_combo = QComboBox()
        self.email_combo.addItems(self.emails)

        template_label = QLabel("Select Template:")
        self.template_combo = QComboBox()
        self.template_combo.addItems(
            list(self.templatesData[x]["meta"]["title"] for x in self.templatesData)
        )
        self.template_combo.currentIndexChanged.connect(self.display_template)
        
        recipient_label = QLabel("Dest Address:")
        self.recipient_input = QLineEdit()
        completer = QCompleter(self.destinations, self.recipient_input)
        self.recipient_input.setCompleter(completer)

        subject_label = QLabel("Subject:")
        self.subject_input = QLineEdit()

        input_layout = QHBoxLayout()
        mailPreviewLayout = QVBoxLayout()
        
        input_layout_field = QVBoxLayout()
        
        


        mail_preview_label = QLabel("Mail Preview:")
        self.mail_preview_edit = QTextEdit()
        self.mail_preview_edit.setReadOnly(True)

        send_button = QPushButton("Send Email")
        send_button.clicked.connect(self.send_email)

        layout.addWidget(email_label)
        layout.addWidget(self.email_combo)
        layout.addWidget(template_label)
        layout.addWidget(self.template_combo)

        layout.addWidget(recipient_label)
        layout.addWidget(self.recipient_input)

        layout.addWidget(subject_label)
        layout.addWidget(self.subject_input)

        mailPreviewLayout.addWidget(mail_preview_label)
        mailPreviewLayout.addWidget(self.mail_preview_edit)

        input_layout.addLayout(mailPreviewLayout)
        input_layout.addLayout(input_layout_field)
        layout.addLayout(input_layout)
        layout.addWidget(send_button)
        self.layout_ = layout
        self.input_layout = input_layout_field
        if len(self.templates) > 0:
            self.current_template = self.templates[0]
            self.display_template(0)

    def on_tab_changed(self, index):
        # Rendre la barre d'outils visible uniquement lorsque vous êtes sur une certaine tab (par exemple, l'index 0)
        self.toolbar.setVisible(index == 1)

    def setup_new_template_tab(self, tab_widget):
        layout = QHBoxLayout(tab_widget)
        subLayout = QVBoxLayout()
        new_template_edit = QTextEdit()
        new_template_edit.setHtml("<p>Enter your HTML template here<p>")

        # Créer une barre d'outils
        self.toolbar = QToolBar("Toolbar", self)
        self.addToolBar(self.toolbar)
        self.toolbar.setVisible(False)
        # Action pour le gras
        bold_action = QAction('Gras', self)
        bold_action.triggered.connect(lambda :self.toggle_bold(new_template_edit))
        self.toolbar.addAction(bold_action)

        # Action pour l'italique
        italic_action = QAction('Italique', self)
        italic_action.triggered.connect(lambda  :self.toggle_italic(new_template_edit))
        self.toolbar.addAction(italic_action)

        # Action pour le souligné
        underline_action = QAction('Souligné', self)
        underline_action.triggered.connect(lambda  : self.toggle_underline(new_template_edit))
        self.toolbar.addAction(underline_action)

        variables_label = QLabel("Template Variables:")
        variables_edit = QTextEdit()
        variables_edit.setReadOnly(True)

        template_name_label = QLabel("Template Name:")
        template_name_edit = QLineEdit()
        template_name_edit.setPlaceholderText("Enter the template name")

        template_title_label = QLabel("Template Title:")
        template_title_edit = QLineEdit()
        template_title_edit.setPlaceholderText("Enter the template title")

        save_template_button = QPushButton("Save Template")
        save_template_button.clicked.connect(
            lambda: self.save_template(new_template_edit, template_name_edit,variables_edit)
        )

        subLayout.addWidget(template_title_label)
        subLayout.addWidget(template_title_edit)
        subLayout.addWidget(template_name_label)
        subLayout.addWidget(template_name_edit)
        subLayout.addWidget(save_template_button)

        layout.addWidget(new_template_edit)
        subLayout.addWidget(variables_label)
        subLayout.addWidget(variables_edit)
        layout.addLayout(subLayout)
        new_template_edit.textChanged.connect(
            lambda: self.update_template_variables(new_template_edit, variables_edit)
        )



    def toggle_bold(self,text_edit):
        print(text_edit)
        self.set_text_format("bold",text_edit)

    def toggle_italic(self,text_edit):
        self.set_text_format("italic",text_edit)

    def toggle_underline(self,text_edit):
        self.set_text_format("underline",text_edit)

    def set_text_format(self, format_name,text_edit):
        # Obtenez le curseur de texte
        cursor = text_edit.textCursor()

        # Obtenez le format actuel
        current_format = cursor.charFormat()

        # Modifiez le format en fonction de l'action
        if format_name == "bold":
            current_format.setFontWeight(2 if current_format.fontWeight() == 50 else 50)  # 2: Bold, 50: Normal
        elif format_name == "italic":
            current_format.setFontItalic(not current_format.fontItalic())
        elif format_name == "underline":
            current_format.setFontUnderline(not current_format.fontUnderline())

        # Appliquez le nouveau format au curseur
        cursor.mergeCharFormat(current_format)

        # Appliquez le curseur au QTextEdit
        text_edit.setTextCursor(cursor)

    def  save_template(self, new_template_edit:QTextEdit, template_name_edit,variables_edit):
        try:
            template_name = template_name_edit.text() + ".html"
            template_content = toHtml(new_template_edit.toPlainText()) 
            # Save the template to the file
            with open(os.path.join(self.templates_dir, template_name), "w") as file:
                file.write(template_content)
            self.templates.append(template_name)
            self.templatesData[template_name] = {"content": template_content}
            # Save metadata to a separate file
            meta_data = {
                "meta": {
                    "title": template_name_edit.text(),
                    "created_at": str(datetime.datetime.now()),
                },
                "var": variables_edit.toPlainText().split("\n")
            }   

            with open(
                os.path.join(self.templates_dir, template_name.replace(".html", ".yaml")),
                "w",
            ) as file:
                file.write(yaml.dump(meta_data))
            self.templatesData[template_name].update(meta_data)
            # Tostify success
            notification_title = "Nouveau template ajouté"
            notification_message = f"Nouveau template {template_name} ajouté avec succès."
            notification.notify(
                title=notification_title,
                message=notification_message,
                app_name='Template Editor'
            )
        except Exception as e:
            # Tostify error
            notification_title = "Erreur lors de l'ajout du template"
            notification_message = f"Erreur lors de l'ajout du template: {e}"
            notification.notify(
                title=notification_title,
                message=notification_message,
                app_name='Template Editor'
            )
            #
            os.remove(os.path.join(self.templates_dir, template_name))
            self.templates.remove(template_name)
            del self.templatesData[template_name]
            raise e
        self.update_template_combo()

    def update_template_combo(self):
        self.template_combo.clear()
        self.template_combo.addItems(
            list(self.templatesData[x]["meta"]["title"] for x in self.templatesData)
        )
        self.layout_.update()

    def setup_edit_confif_tab(self, tab_widget):
        layout = QVBoxLayout(tab_widget)

        smtp_config_label = QLabel("SMTP Configuration:")
        smtp_config_edit = QTextEdit()
        smtp_config_edit.setPlainText(yaml.dump(self.smtp_config))
        
        save_config_button = QPushButton("Save Configuration")
        save_config_button.clicked.connect(lambda: self.save_config(smtp_config_edit))

        layout.addWidget(save_config_button)
        layout.addWidget(smtp_config_label)
        layout.addWidget(smtp_config_edit)
    
    def save_config(self, smtp_config_edit):
        try:
            save = ""
            with open(self.config_file, "r") as file:
                save = file.read()
            with open(self.config_file, "w") as file:
                file.write(smtp_config_edit.toPlainText())
            self.update_smtp_config()
            # Tostify success
            notification_title = "Nouvelle configuration ajoutée"
            notification_message = f"Nouveaux paramètres ajoutés avec succès."
            notification.notify(
                title=notification_title,
                message=notification_message,
                app_name='Config Editor'
            )
        except Exception as e:
            # Tostify error
            notification_title = "Erreur lors de l'ajout de la configuration"
            notification_message = f"Erreur lors de l'ajout de la configuration: {e}"
            notification.notify(
                title=notification_title,
                message=notification_message,
                app_name='Config Editor'
            )
            #
            with open(self.config_file, "w") as file:
                file.write(save)
            self.update_smtp_config()
            

    def update_smtp_config(self):
        smtp_config = []
        with open(self.config_file, "r") as file:
            smtp_config = yaml.safe_load(file)
        self.emails = []
        print(smtp_config)
        for k in smtp_config:
            print(k)
            self.emails.append(k + "=>" + smtp_config[k]["user"])
        self.email_combo.clear()
        self.email_combo.addItems(self.emails)
        self.layout_.update()
        self.smtp_config = smtp_config

    def update_template_variables(self, template_edit, variables_edit):
        try:
            template_content = template_edit.toPlainText()
            env = Environment(loader=FileSystemLoader(self.templates_dir))
            rendered_template = env.parse(template_content)
            variables = meta.find_undeclared_variables(rendered_template)
            variables_edit.setPlainText("\n".join(variables))
        except Exception as e:
            pass


    def display_template(self, index):
        template = self.templates[index]
        self.current_template = template
        env = Environment(loader=FileSystemLoader(self.templates_dir))
        rendered_template = env.parse(self.templatesData[template]["content"])
        self.mail_preview_edit.setPlainText(extractBody(self.templatesData[template]["content"]))
        self.clear_var_inputs()
        word_list = list("!"+x[2:]+"()" for x in globals() if x.startswith("e_") )
        var_layout = QVBoxLayout()  # Utiliser un QVBoxLayout pour contenir les champs variables
        for var in meta.find_undeclared_variables(rendered_template):
            var_label = QLabel(var + ":")
            var_input = QPlainTextEdit()
            var_input.setPlaceholderText("Enter " + var)
            # Bind the variable to the input field
            var_input.textChanged.connect(lambda : self.update_mail_preview(var))
            self.vars[var] = var_input
            var_input.setFixedHeight(30)
            completer = CustomCompleter(word_list,var_input)
            completer.setWidget(var_input)
            completer.setCompletionMode(QCompleter.PopupCompletion)
            completer.setCaseSensitivity(Qt.CaseInsensitive)
            var_input.textChanged.connect(lambda c=completer,v= var_input: self.autocomplete_text(c, v))
            self.var_inputs[var] = var_input
            self.var_inputs[var+"_"] = var_label
            
            var_layout.addWidget(var_label)
            var_layout.addWidget(var_input)

        # Ajouter le layout des champs variables à l'interface
        self.input_layout.addLayout(var_layout)

    def autocomplete_text(self,completer,editor):
        cursor =editor.textCursor()
        cursor.movePosition(QTextCursor.StartOfBlock, QTextCursor.KeepAnchor)
        selected_text = cursor.selectedText().strip()
        cursor_pos = editor.mapFromGlobal(editor.cursorRect().bottomRight())
        completer.popup().move(editor.mapToGlobal(cursor_pos))
        completer.setCompletionPrefix(selected_text)
        completer.complete()

    def update_mail_preview(self, var):
        template = self.current_template
        env = Environment(loader=FileSystemLoader(self.templates_dir))
        rendered_template = env.from_string(self.templatesData[template]["content"]).render(
            **{var: self.parseF(self.vars[var].toPlainText(),var) for var in self.vars}
        )
        self.mail_preview_edit.setPlainText(extractBody(rendered_template))
    def parseF(self,val,key):
        if val.startswith("!"):
            try:
                return eval("e_"+val[1:],globals())
            except Exception as e:
                return val
        elif len(val) == 0:
            return "<"+key+">"
        else:
            return val
    def clear_var_inputs(self):
        self.vars = {}
        for _,var_input in self.var_inputs.items():
                self.input_layout.removeWidget(var_input)
                var_input.deleteLater()
        self.var_inputs = {}

    def send_email(self):
        email = self.email_combo.currentText().split("=>")[0]
        email_data = self.smtp_config[email]

        recipient = self.recipient_input.text()
        subject = self.subject_input.text()
        template_file =  self.current_template
        env = Environment(loader=FileSystemLoader(self.templates_dir,encoding='latin1' ))
        rendered_template = env.get_template(template_file).render(self.get_var_values())
        print(f"""
Recap : 
    sender : {email}
    dest : {recipient}
    object : {subject}
{rendered_template}   
""")

        # Code to send the email using SMTP configuration
        sender_email = email_data['user']
        password = email_data['password']
        smtp_server = email_data['smtp']
        smtp_port = email_data['port']
        # Configurer le message
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient
        msg['Subject'] = subject
        msg.attach(MIMEText(rendered_template, 'html'))
        try:
            # Établir la connexion au serveur SMTP
            with smtplib.SMTP(smtp_server, smtp_port) as server:
                server.starttls()
                # Login au compte SMTP
                server.login(sender_email, password)
                # Envoyer l'email
                server.sendmail(sender_email, recipient, msg.as_string())
            # Notify success
            notification_title = "Email envoyé"
            notification_message = f"Email envoyé avec succès à {recipient}."
            notification.notify(
                title=notification_title,
                message=notification_message,
                app_name='Email Sender'
            )
            # Clear the fields
            self.recipient_input.setText("")
            self.subject_input.setText("")
            self.mail_preview_edit.setText("")
            self.clear_var_inputs()
            # Update Sender List

            if not recipient in self.destinations:
                with open(self.dest_file, "a") as file:
                    file.write("\n"+recipient + "\n")
                self.destinations.append(recipient)
            # Update Recipient List on GUI
            completer = QCompleter(self.destinations, self.recipient_input)
            self.recipient_input.setCompleter(completer)
            self.layout_.update()
        except Exception as e:
            # Notify error
            notification_title = "Erreur lors de l'envoi de l'email"
            notification_message = f"Erreur lors de l'envoi de l'email: {e}"
            notification.notify(
                title=notification_title,
                message=notification_message,
                app_name='Email Sender'
            )
            raise e
    def get_var_values(self):
        var_values = {}
        for k,var_input in self.var_inputs.items():
            if isinstance(var_input, QLabel):
                continue
            var_value = self.parseF(var_input.toPlainText(),f"Missing Var : {k}")
            var_values[k] = var_value
        return var_values


if __name__ == "__main__":
    app = QApplication(sys.argv)
    email_app = EmailApp()
    email_app.show()
    sys.exit(app.exec_())
