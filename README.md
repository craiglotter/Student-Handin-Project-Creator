Student Handin Project Creator
==============================

Commerce I.T. has a special network folder with restrictive file access policies in place to handle student project hand-ins. Student Handin Project Creator is an application to automate the process of creating specific hand-in folders for specified projects on lecturer request. It works by generating the specified main project folders and then automatically retrieving the list of registered students for that course/project, generating individual subfolders and assigning exclusive rights to each individual folder using a JRBUtils console application. It also generates an access.ini file that control student hand-in times when using "Student Handin System" from the student computer lab.

Note: Created for Commerce I.T.
Note: Update: Handin shifted from a Novell server to a Windows 2003 server.

Created by Craig Lotter, January 2006

*********************************

Project Details:

Coded in Visual Basic .NET using Visual Studio .NET 2003 (Upgraded to .NET 2005)
Implements concepts such as Folder manipulation, Shell scripting.
Level of Complexity: Simple

*********************************

Update 20070328.05:

- Upgraded to a Studio .NET 2005 project
- With the decommissioning of the Novell based Comlab file server, it was decided to move the handin functionality to the online sphere, namely the Commerce webserver. This is a Windows 2003 server so the Project Creator application needed to be updated accordingly.
- Now allows the user to specify creating student folders for a Novell or Windows environment. (Note: Windows environment - folder rights are determined by the parent folder rights.

*********************************

Update 20070813.06:

- Added a batch handin folder creation option.
- Added Help and About dialogs

*********************************

Update 20080212.07:

- Fixed bug in that the batch create option was not generating the access.ini file
- Updated the application look slightly.
