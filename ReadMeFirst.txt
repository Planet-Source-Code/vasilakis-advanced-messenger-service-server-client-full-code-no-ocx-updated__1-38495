PLEASE READ THIS DOCUMENT BEFORE ANYTHING!

This is a messenger service created just for fun. It is very nice written, and *as much as I could* bug-free. It includes: the Messenger client, the messenger Server, the Banner Server and the database of the program for SQL Server. It support multi-file transfer and many other features. It is just like MSN Messenger but of course not that advanced :P You should see it, it might help you a lot. Everything is pure source code. That means you need a lot of api calls :p  The projects does not have comments, but I think it won't be very difficult to understand my functions and subs. Please send me a feedback, or email me for any questions at vasilis@vasilakis.com.

Requirements:

MS VB 6.0 + some components as winsock, common controls etc.
MS SQL Server (at least 7.0) OR Microsoft Access Database Components
MS ADO 2.6

------------------------------------------------------------------------------

* To make Messenger & Server work:
	*FOR SQL:
		Create a database named "vsag" @ your SQL Server.
		Take "sqlDB" File and restore it as backup @ your SQL Server 
		as database name "vsag".
		Open the "Server" folder and check the COnnection String for the SQL
		Server at module "modTasks" Sub Main(). When done, 
		execute the project.
	*FOR Access:
		Copy the mdbAccess.mdb at the Server folder. Change the connection string
		at module "modTasks" Sub Main(). When done, execute the project.

	There is a user added with username "vasilis@vasilakis.com" 
	and password 123 (or something, check it at table "users").
	NOW! Open the Client folder and open the project messenger.
	Execute it 
	Login (for server type 'localhost' or the name of your PC).

* To make BannerServer work:
	BEFORE ANYTHING: Open the Table "banners" 
	@ database "vsag" on your SQLServer OR the mdbAccess.mdb and CHECK THE 
	PATHS OF THE IMAGES!!!
	Open the Banner folder and execute the prjBanner.vbp.

	Reconnect the messenger client.

------------------------------------------------------------------------------
Enjoy!!!

For comments:
vasilis@vasilakis.com

Please visit my Web URL:
www.vasilakis.com