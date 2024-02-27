# Fast Mail Sender

## Index

- [Install](#install)
- [Usage](#usage)
- [Who to use](#who-to-use)
  - [First panel](#first-panel)
  - [Second panel](#second-panel)
  - [Third panel](#third-panel)
  - [Type of template](#type-of-template)

# Install 

Run the following command to install the package:

```bash
pip install -r requirements.txt
```

# Usage

```bash
python main.py
```

# Who to use

You got 3 panel : 
- The first one is for the sender email
- The second one is for template creation
- The third one is for e-mail configuration

## First panel

You can send an email from one of your configured email. [here](#third-panel)
Select the email you want to use.
Select the email template you want to use [here](#second-panel) 
Enter the destination email(It's store destination email in a file for future use)
Enter the subject of the email
Complete template variables
Click on send button

**Templates variables**
You can insert simple text in variable, or use defined function to generate a value.
The function must be in the following format : 
```python
def e_<func_name>():
    return <value>
```
and then use it in the template like this : 
```python
!<func_name>()
```

> Auto completion is available for the function name

## Second panel

You can create a new template. Use HTML and CSS to beautify your email.
To declare a variable, use the following format : 
```html
{{<variable_name>}}
```

Template Title : The title of the template in the selected
Template Name : The name of the template files

## Third panel

You can configure your email to send email from it.
It's use YAML syntax to store the email configuration.
You can add multiple email configuration.

```yaml
<email_name>:
    user: <email>
    password: <password>
    smtp: <smtp>
    port: <port>
```


## Type of template

Here an example of a template : 
```html

<h1> Hello {{name}} </h1>

<p> This is a test email </p>

<p> Your age is : {{age}} </p>

Regards,
{{signature}}
```

If y have set you signature in `main.py:33` you can fill the variable `signature` in the first panel, with it : `!signature()`. 

