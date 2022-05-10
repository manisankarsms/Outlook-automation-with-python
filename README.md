# Outlook-automation-with-python

Triggering a python script when a email with specific subject or from specific mail

- create a outlook macro (VBA) to trigger the python program
  Script:
		Sub run_python(item As Outlook.MailItem)
		emailBody = item.Body
		Shell ("python C:\Users\manis\Desktop\PythonAuto\python_code.py" & " " & emailBody)
		End Sub
 
- create a outlook rule to trigger the above created script based on the required conditions

- in most of the outlook versions, run a script in rules was disabled, we need to enable it from windows registry

  
