# Payroll--system
This is a Basic payroll system which calculates your gross salary and gives you the net pay design with Python

Project Overview 
The system allows a user to audit a payroll for the employees without worries. 

The Features
.Automation of Taxes - The system calculates the employees Net Pay automatically
.Automation of PDF files individually - The system generates PDF's for the employees automatically saving time in the process
.Auto send System - The system reads the  employees phone numbers and automatically sends them to whatsapp messaging application.There is a system whereby it can also automatically send emails. If needed I customize for the user individually to enable that feature. 

The Goal of this Project is to help  businesses which would like to grow efficiently , transcend to evolving technological world and save time on the whole tax auditing  shenanigans.

Tips: On the screenshot folder you can see how the Payroll system works it is well documented and the python scripts are there 
The first image starts at number 3.png to 18.png which is the last image 

Here's the python script just incase you miss them: 

How to install the dependencies for the python code 
on debian based / ubuntu based linux distros on windows you can try running it on vs code terminal 

Sudo apt install python3-pandas python3-openpyxl -y 

To run the python code run on the  vs code terminal 

python3 payroll.py create 

This creates the excel document 

To sync the excel document after  editing run on vs code terminal

python3 payroll.py sync 

To automatically send the PDF to the phone number on whatsapp run on vs code terminal 

python3 payroll.py whatsapp


