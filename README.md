# Access Query Runner
### A utility for enumerating and executing queries (procedures and views) from Microsoft Access.

## Overview
If you find yourself obligated for whatever reason to use Microsoft Access as your "backend" database for a project, you'll encounter many difficulties. If you search online for solutions or guidance in dealing with these difficulties, the general solution offered by everyone will be to avoid using Microsoft Access in your projects and it will be left at that. It will be like asking people to help you minimize the damage to your merchandise as you deliver fine china by pogo stick. No one will have any advice on minimizing the damage, but everyone will tell you to abandon your pogo stick.

Of course. Everyone knows that Microsoft Access and pogo stick deliveries are bad ideas. Even the people at Microsoft.

But, depending on the client, it may simply not be possible to prevent the use of Access. You'll be stuck with its problems and few people will be willing to help.

## A Little Help in a Small Corner
One problem you might encounter is in providing an interface for users to be able to execute "queries" (i.e., views and procedures) that they've created or that they might create in Access. It's easy enough to develop an interface to run specific queries that you already know about or that you have written yourself. The difficulties arise if you allow users to create their own views and procedures in Access and you try to "snoop" the database to enumerate those views and procedures and execute them.

Among other peculiarities in Access, it seems that **views** may have names with spaces, but **procedures** cannot. The problem is that within Access you don't really know if you're creating a view or a procedure. If you specify a parameter in the *Design View* of a query in Access, you will be creating a procedure. If you don't specify a parameter, you'll create a view. And you **can** use spaces in Access itself for the name of a procedure but then you won't be able to execute that procedure from C# if you invoke that procedure by name. In order to execute views and procedures that have names with spaces in them, you have to "wrap" such names with brackets before you invoke them with such methods as the `ExecuteReader` method of the `OleDbCommand` object.

Certain Access functions (such as **Nz** and **InStrRev**) work only within the Access environment. You can't execute procedures or views containing such expressions through the standard OLEDB driver. If such a query is identified by the Query Runner, it is enumerated but it can't be selected as a query to execute.

Also note that not all information for views, procedures, and parameters is available for inspection and retrieval using the methods in the `System.Data.OleDb` namespace. Specifically, the `GetOleDbSchemaTable` method of the `OleDbConnection` object doesn't support passing in the 'Procedure_Parameters' GUID in order to retrieve parameter information for procedures in an Access database. If you attempt to pass the 'Procedure_Parameters' GUID as a parameter to the `GetOleDbSchemaTable` method, you'll generate an exception along the lines of the following:

> The Procedure_Parameters OleDbSchemaGuid is not a supported schema by the 'Microsoft.ACE.OLEDB.12.0' provider.'

So, I've had to resort to the old `Microsoft.Office.Interop.Access.Dao` stuff in order to inspect the Access database for relevant information with respect to procedure parameters. The DAO library is used in this project only to retrieve the parameter information in the `GetQueryList` method of the `DataService` class. In the `GetResultsTable` method, the `System.Data.OleDb` methods are used to retrieve actual data from the database using those parameters.

This project may be of no use to anyone, but it's one approach to tackling the problem of running queries in Access in a more or less flexible way. This project doesn't necessarily represent the best way to go about this, but if you need a solution along these lines for a project based on Access, you could do worse.

## The Project
**Access Query Runner** is a simple C# WPF solution developed in Visual Studio 2017. The solution as it is works with Microsoft Access 2007 or above (that is, versions of Access that use the "Microsoft ACE OLEDB 12.0" database engine), but you can modify the database connection string for earlier versions of Access. (And you might need to make consequent modifications to references, e.g., you may have to use the Microsoft DAO 3.6 Object Library instead of the Access interop assemblies. I can't claim to have tested the project with all versions of Access.)

The **Access Query Runner** enumerates procedures and views in an Access database and allows the user to execute those procedures and views, outputting results to text files or to Excel.

In addition to the Query Runner project, a couple of sample Access databases are included.

## Credits
Thanks are due to a [Stack Overflow contributor](https://stackoverflow.com/users/4048/compile-this) who posted a [response](https://stackoverflow.com/a/38064) for deriving start and end of week DateTime values using extension methods. I've adapted those methods to add some convenience utilities to this project.

## License
MIT License. See [LICENSE](LICENSE) for more information.