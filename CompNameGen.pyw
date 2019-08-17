# CompNameGen.pyw 
# This creates a GUI for a powershell script  CreateComputerNames.ps1

from tkinter import *
import subprocess, sys
import os
import ctypes


def MakeGUI():
	OUInput=e3.get()[1]
	if e1.get()=="":
		ctypes.windll.user32.MessageBoxW(0, u"Please input the first part of the Computer Name", u"Invalid Input Error", 0) #Create an error message box
	elif (e2.get()=="" or (e2.get().isdigit())==False):
		ctypes.windll.user32.MessageBoxW(0, u"Please input the Total Amount of Computers in positive Integers Only", u"Invalid Input Error!", 0)
	elif (e3.get()=="" or (e3.get()=='""') or(OUInput.upper()!="O")): #Need to create better error hanlding
		ctypes.windll.user32.MessageBoxW(0, u"Please input the full path to OU surronded by quotes", u"Invalid Input Error", 0)
	else:
		ctypes.windll.user32.MessageBoxW(0, u"This may take a few minutes.  Please be patient!", u"Processing please wait...",0)
		powershellpath= r'C:\windows\System32\WindowsPowerShell\v1.0\powershell.exe'
		powershellCMD= "C:\Scripts\CreateComputerNames.ps1"
		
		p= subprocess.Popen([powershellpath, '-ExecutionPolicy', 'Unrestricted', powershellCMD, e1.get(), e2.get(), e3.get()], stdout=subprocess.PIPE, stderr=subprocess.PIPE) 
		output, error = p.communicate()
		rc = p.returncode
		print ("Return code to Python is: " + str(rc))
		print ("\n\nstdout: \n\n" + str(output))
		print ("\n\nstderr: " + str(error))
		strError=str(error)
		if strError!="b''":
			ctypes.windll.user32.MessageBoxW(0, "Please enter the correct OU in the format of: OU=myou,DC=domain,DC=com for the full OU path.",u"Invalid Entry OU not found!",0) #send error to msgbox
		else:
			ctypes.windll.user32.MessageBoxW(0, u"Computers accounts are now created", u"Operation Complete!",0)
		
def masterquit():
  quit()


master = Tk()
Label(master, text="First part of the computer names: ", bd="2", font=("helvetica", "12", "bold")).grid(row=0,ipadx=3,sticky=W)
Label(master, text="Total amount of computers: ", bd="2", font=("helvetica", "12", "bold")).grid(row=1,ipadx=3,sticky=W)
Label(master, text="Full OU path in quotes: ", bd="2", font=("helvetica", "12", "bold")).grid(row=2,ipadx=3,sticky=W) 

e1 = Entry(master)
e1.grid(row=0, column=1,ipadx=25,sticky=W)
e2 = Entry(master)
e2.grid(row=1, column=1,ipadx=25,sticky=W)
e3 = Entry(master)
e3.grid(row=2, column=1,ipadx=255,sticky=W)
e3.insert(10, '""')


Button(master, text='QUIT', fg="red", bd="8", font=("helvetica", "14", "bold"),activebackground="orange", width="20", relief="raised", command=masterquit).grid(row=3, column=0, sticky=W, pady=8, padx=4)
Button(master, text='Create Computer Names', fg="green", bd="8", font=("helvetica", "14", "bold"),activebackground="orange", width="20", relief="raised", command=MakeGUI).grid(row=3, column=1, sticky=W, pady=8, padx=4)

master.title("Generate New Computer Names")
master.geometry("925x250+450+450")
mainloop( )
