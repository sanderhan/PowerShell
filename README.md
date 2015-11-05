# PowerShell Module to SQL Server Data Import/Export
Collection of useful PowerShell scripts and modules

<# 
 .Synopsis
  PNI Data Utility.

 .Description
    Useful functions to export/import data from database.

    To import this module, use the following commands in script

        Push-Location
            Import-Module sqlps -DisableNameChecking
            Import-Module <your-modules-location>\SQLDataUtility.psm1
        Pop-Location

    To Unload Module, use
        Remove-Module SQLDataUtility


 .Functions
    
    Invoke-BCP -OpType <String> {IN | OUT | QUERYOUT } –DataFile <String> -Server <String> -Database <String> -Table <String> -Query <String> [-User <String>] [-Password <String>] [ –Format <String> {Native | CSV} ]
    Export-Tables  –DataPath <String> -Server <String> -Database <String>  [-User <String>] [-Password <String>] [–Tables <String[]>] [-SchemaIncludes <String[]>] [-SchemaExcludes <String[]>] [-TableIncludes <String[]>] [-TableExcludes <String[]>]
    Import-Tables  –DataPath <String> -Server <String> -Database <String>  [-User <String>] [-Password <String>] [–Tables <String[]>] [-Clear]
    Email-Directory  –Path <String> -To <String[]> -Subject <String>  [-Message <String>]
    Email-Files  –Files <String[]> -To <String[]> -Subject <String>  [-Message <String>] [-Compress] [-CompressFileName <String>] 

    
 .Example
    # Export data from tables in ,<atabase> on Server <db instance> to folder C:\Dump, table <schema>.<mytable> is not exported.
    Export-Tables -Server >db instance> -Database <mydb> -TableExcludes "schema.mytable" -DataPath C:\Dump

 .Example
    # Email all files under C:\Dump
    Email-Directory -Path C:\Dump -To "xxx@xxxxx.com" -Subject "zipped data file." -Message "Data file is attached."  

 .Example
    # Import data from data files which were exported by export-tables under C:\ImportData
    Import-Tables -Server localhost -Database <mydb> -DataPath C:\ImportData -Clear

 
#>
