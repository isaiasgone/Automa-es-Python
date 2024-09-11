import win32com.client

sapguiauto = win32com.client.GetObject('SAPGUI')
application = sapguiauto.GetScriptingEngine
connection = application.Children(0)
session = connection.children(0)

print(type(session))