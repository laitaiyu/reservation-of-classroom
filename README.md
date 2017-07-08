# reservation-of-classroom
How to establish a reservation of classroom web system using ASP.NET?

# Introduction

This is a reservation of classroom web system, but you can change it to make another reservation of the things web system. Such as professional classroom, meeting room or hotel room. Its operation is very easy and design or redesign is very simple. This is an open source code of the web system, you can free use and change it. No pay for me. I will teach you how to operate and redesign it.
The roles are an administrator and many teachers. The administrator manages this web and the teacher is reservation classroom of the operator. We need the administrator create many classroom’s name, item’s name, context (subject), teacher’s name and office’s name. He can query all reservations of whole teacher status. The teachers are able to reserve a classroom. They can choose which classroom, which context, which office, which item and which time and they also are able to review reversation status.
The web system is very small and the database is using access 2003 format. You also are able to using Microsoft Access to edit it. There are two interfaces that you can manage the database. Firstly, the web system provided a management interface. Secondly, To use Microsoft Access, manage and edit the database.
Because it is using Microsoft.Jet.OLEDB.4.0, if you install the Internet Information Services (IIS) 7, you should enable 32 bit application.


# Background

Equipment
Operation System: Microsoft Windows 7 (64 bit)
Web System: Internet Information Services (IIS) 7
Development Utility: Microsoft Visual Studio 2010


# Using the code


1. Copy this web system to your hard disk. 

1.1. To extract the RAR file to your hard disk and copy the folder to under the folder called ‘wwwroot’ that is the root folder of IIS 7.

1.2. To rename the folder name to ‘WebApplication1’.

1.3. To convert the ‘WebApplication1’ to application.

1.4. Enable 32 bit application, it needs to setup that is true.


2. This is a homepage of mainly. The default administrator is ‘admin’ and default password is ‘admin’. The default teacher login name is ‘1’ and default password is ‘1’.
You can choice a classroom and a date and review the reservation status.

3. This is the admin utility. You can see those topics, including classroom, context, office, teacher, item and reservation and you are able to manage the classroom, the context, the office, the teacher and the item.

4. You can use the same interface to add, update, delete, edit and cancel those items, such as the classroom, the context, the office, the teacher and the item.

5. The reservation provided review status of reservations.

6. If you are using a teacher account to login this web system, you should reserve to a classroom.

