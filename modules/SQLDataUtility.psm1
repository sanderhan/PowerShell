<# 
 .Synopsis
  SQL Server Data Utility for Import/Export data.

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

 .Example
    # Export data from tables in ,<atabase> on Server <db instance> to folder C:\Dump, table <schema>.<mytable> is not exported.
    Export-Tables -Server >db instance> -Database <mydb> -TableExcludes "schema.mytable" -DataPath C:\Dump

 .Example
    # Import data from data files which were exported by export-tables under C:\ImportData
    Import-Tables -Server localhost -Database <mydb> -DataPath C:\ImportData -Clear

 
#>

function Format-output ([String] $Message, 
                        [ValidateSet("INFO", "DEBUG", "ERROR")] [String] $Type = "INFO") {
    $Message = "[{0:yyyy-MM-dd HH:mm:ss.fff}]:[{1,5}]:{2}" -f (Get-Date) , $Type, $Message
    $Message
}
 

function Invoke-BCP([ValidateSet("IN", "OUT","QUERYOUT")] [String]$OpType, 
                    [String]$DataFile, 
                    [String]$Server,
                    [String]$Database, 
                    [String]$Table, 
                    [String]$Query,
                    [String]$User, 
                    [String]$Password,
                    [ValidateSet("Native", "CSV")] [String]$format = 'Native' ){

    Format-output "BCP $Table $OpType ..." |Write-Output

    If($OpType -eq 'QUERYOUT' ){
        $source = $Query
    }else{
        $source = $Table
    }
    
    If($format -eq 'CSV'){
        $formatOption = @('-w',"-t,")
    }else{
        $formatOption = '-N'
    } 

    if($User) {

        [String[]] $bcpArgs = @("$source", "$OpType" , "$DataFile", "-E","-k",
                                "-d", "$Database",
                                "-S", "$Server",
                                "-U", "$User",
                                "-P", "$Password")

    }else{
        [String[]] $bcpArgs = @("$source", "$OpType" , "$DataFile", "-E","-k","-T",
                                "-d", "$Database",
                                "-S", "$Server")
    }

   $bcpArgs  += $formatOption

   $output = [string](& "bcp.exe" $bcpArgs)
   $outputMsg = $output -join "`n"

   if( $lastexitcode -ne 0){ 
        Throw "BCP $source $OpType failed."
   }
   Write-Verbose -Message $outputMsg

   if( Select-String -InputObject $outputMsg -AllMatches "ERROR" ){ 
        Format-output -Message "BCP $source $OpType failed due to error action." -Type ERROR | Write-Error 
        Throw "BCP $source $OpType failed due to error action."
   }

   Format-output "BCP $source $OpType is completed." |Write-Output
} 


function Clear-Tables([String[]] $Tables, 
                       [String] $Server,  
                       [String]$Database, 
                       [String]$User, 
                       [String]$Password){

    
    $sortQuery = ";WITH a 
                   AS 
                   ( SELECT 0 AS lvl, t.object_id AS tblID 
                     FROM sys.TABLES t 
                     WHERE t.is_ms_shipped=0 
                     AND t.object_id NOT IN (SELECT f.referenced_object_id FROM sys.foreign_keys f) 
                     UNION ALL 
                     SELECT a.lvl + 1 AS lvl, f.referenced_object_id AS tblId 
                     FROM a 
                        INNER JOIN sys.foreign_keys f  ON a.tblId=f.parent_object_id AND a.tblID<>f.referenced_object_id ) 
                   SELECT '['+ object_schema_name(tblID) + '].[' + object_name(tblId) + ']' as TableName FROM a
                   WHERE tblID in ( "; 



    $sbSort = New-Object -TypeName System.Text.StringBuilder
    $sbDel = New-Object -TypeName System.Text.StringBuilder
    $null = $sbSort.Append($sortQuery);
      

     ForEach($table in $tables){
        
        $inList = ("OBJECT_ID('{0}')," -f $table)
        $null = $sbSort.Append($inList)
        $null = $sbDel.Append("DELETE FROM $table;`n")

     }
     $null = $sbSort.Append("-1) `n GROUP BY tblId ORDER BY MAX(lvl),1 ")
     $null = $sbDel.Append("COMMIT TRAN;`n")
     $sortSql = $sbSort.ToString()

     if($User){
        $sortTables = Invoke-SQlcmd -ServerInstance $Server -Database $Database -Query $sortSql -Username $User -Password $Password
     }else{
        $sortTables = Invoke-SQlcmd -ServerInstance $Server -Database $Database -Query $sortSql 
     }

     $sbDel = New-Object -TypeName System.Text.StringBuilder
     $null = $sbDel.Append("SET XACT_ABORT ON;`nBEGIN TRAN;`n");
     ForEach($row in $sortTables){
        $tableName = $row.TableName;       
        $null = $sbDel.Append("DELETE FROM $tableName;`n")
     }
     $null = $sbDel.Append("COMMIT TRAN;`n")

     $delSql = $sbDel.ToString()

     if($User){
        Invoke-SQlcmd -ServerInstance $Server -Database $Database -Query $delSql -Username $User -Password $Password
     }else{
        Invoke-SQlcmd -ServerInstance $Server -Database $Database -Query $delSql
     }
     
     Format-output "Tables have been deleted in database $Database on Server $Server." |Write-Output
} 

function Export-Tables([String[]] $Tables, 
                       [String] $Server,  
                       [String]$Database, 
                       [String]$User, 
                       [String]$Password,
                       [String[]]$SchemaIncludes ,
                       [String[]]$SchemaExcludes,
                       [String[]]$TableIncludes ,
                       [String[]]$TableExcludes,
                       [String]$DataPath){

                                                                                                                                                       $dataFileLocation = $DataPath 
    $zipped = $false
    $dataFileLocation = $DataPath

    If($DataPath -match "\.zip$"){
        $zipFileName = (Split-Path -Path $DataPath -Leaf) -replace "\.zip$"

        $dataFileLocation = Join-Path -Path $env:TEMP -ChildPath ("{0}_{1}" -f $zipFileName, (Get-Date -format "yyyyMMddmmHHss") )  

        $zipped = $true
    }


    [String[]] $outFailedTables = @()

    If(!$Tables){
        $Tables = Get-Tables -Server $Server -Database $Database -User $User -Password $Password `
                             -SchemaIncludes $SchemaIncludes -SchemaExcludes $SchemaExcludes -TableIncludes $TableIncludes -TableExcludes $TableExcludes
    }

    If(!(Test-Path -Path $dataFileLocation) ){
        New-Item $dataFileLocation -ItemType Directory -Force |Out-Null
    }

    ForEach($table in $Tables){
       $dataFileName = "{0}.dat" -f  $table 
       $dataFile = Join-Path -Path $dataFileLocation -ChildPath $dataFileName
       try{
            Invoke-BCP -OpType "OUT" -DataFile $dataFile -Server $Server -Database $Database -Table $table -User $User -Password $Password
            Format-output  "Table $table has been exported." |Write-Output
       }catch{
         $outFailedTables += $table;
         Format-output  "Table $table failed to export." -Type ERROR |Write-Error
       }
    }
    if($zipped){
        Compress-Folder -Path $dataFileLocation -Outfile $DataPath
        Remove-Item -Path $dataFileLocation -Force -Recurse|Out-Null
    }

    return $outFailedTables;
}

function Import-Tables([String[]] $Tables, 
                       [String]$Server,  
                       [String]$Database, 
                       [String]$User, 
                       [String]$Password,
                       [String]$DataPath,
                       [Switch]$Clear){

    [String[]] $inFailedTables = @()

    $dataFileLocation = $DataPath 
    $zipped = $false

    If($DataPath -match "\.zip$"){
        $zipFileName = (Split-Path -Path $DataPath -Leaf) -replace "\.zip$"

        $dataFileLocation = Join-Path -Path $env:TEMP -ChildPath ("{0}_{1}" -f $zipFileName, (Get-Date -format "yyyyMMddmmHHss") )  

        Extract-Zipfile -ZipFile $DataPath -ExtractPath $dataFileLocation

        $zipped = $true
    }

    If(!$Tables){
        $Tables = Get-ChildItem -Path $dataFileLocation -Filter "*.dat"  -File | % {$_.Name.TrimEnd(".dat")} 
    }

    If($Clear){
        Clear-Tables -Server $Server -Database $Database -User $User -Password $Password -Tables $Tables
    }

    ForEach($table in $Tables){
          Write-Output "Importing table $table ..."
          try{
              $dataFileName = "{0}.dat" -f  $table 
              $dataFile = Join-Path -Path $dataFileLocation -ChildPath $dataFileName
              Invoke-BCP -OpType "IN" -DataFile $dataFile -Server $Server -Database $Database -Table $table -User $User -Password $Password
              Format-output  "Table $table has been imported." | Write-Output
           }catch{
             $inFailedTables += $table;
             Format-output  "Table $table failed to be imported." -Type ERROR  | Write-Error
           }
    }
 
    If($zipped){
        Remove-Item -Path $dataFileLocation -Recurse -Force|Out-Null
    }
    return $inFailedTables;
}


function Copy-Tables([String[]] $Tables, 
                         [String]$SrcServer,  [String]$SrcDatabase, [String]$SrcUser, [String]$SrcPassword,
                         [String]$DestServer, [String]$DestDatabase, [String]$DestUser, [String]$DestPassword,
                         [String]$DataPath,
                         [Switch]$ExportOnly){

    [String[]] $outFailedTables = @()
    [String[]] $inFailedTables = @()

    [String] $timeExport = (Get-Date -Format "yyyyMMdd_HHmm")

    $dataFilesPath =Join-Path -Path (Join-Path -Path $DataPath -ChildPath $SrcDatabase) -ChildPath $timeExport

    If(!(Test-Path -Path $dataFilesPath) ){
        New-Item $dataFilesPath -ItemType Directory -Force |Out-Null
    }
    
    $outFailedTables = Export-Tables -Tables $Tables -Server $SrcServer -Database $SrcDatabase -User $SrcUser -Password $SrcPassword -DataPath $dataFilesPath
    if(!$ExportOnly){
        $inFailedTables = Import-Tables -Tables $Tables -Server $DesServer -Database $DestDatabase -User $DestUser -Password $DestPassword -DataPath $dataFilesPath -Clear
    }

    $mailErrorMessage ="BCP Failed while transferring data from database $SrcDatabase on $SrcServer to $DestDatabase on $DestServer. `n" 
    $hasError = $false
    if($outFailedTables.Count -gt 0){
       $hasError = $true
       $ErrorMessage += ("BCP Out failed on tables:`n " + ($outFailedTables -join "`n") +"`n")
    }

    if($inFailedTables.Count -gt 0){
       $hasError = $true
       $ErrorMessage += ("BCP In failed on tables:`n " + ($inFailedTables -join "`n") +"`n")
    }
    If($hasError){
        Format-Output $ErrorMessage -Type ERROR  |Write-Error 
    }

    return $hasError

}

function Extract-Zipfile ([String]$ZipFile, [String]$ExtractPath){
    Add-Type -assembly "system.io.compression.filesystem"
    [io.compression.zipfile]::ExtractToDirectory($ZipFile, $ExtractPath) 
}


function Compress-Folder ([String]$Path, [String]$Outfile){
    Add-Type -assembly "system.io.compression.filesystem"
    If(Test-path $Outfile) {Remove-item $Outfile}
    [io.compression.zipfile]::CreateFromDirectory($Path, $OutFile) 
}


function Get-Tables ([String] $Server, [String]$Database,[String]$User,[String]$Password,
                     [String[]]$SchemaIncludes ,[String[]] $SchemaExcludes,
                     [String[]]$TableIncludes ,[String[]] $TableExcludes){
    $tables = @()

    $sb = New-Object -TypeName System.Text.StringBuilder
    $null = $sb.Append("SELECT  '['+ object_schema_name(object_id) + '].[' + name + ']' as TableName FROM sys.tables WHERE is_ms_shipped = 0 `n ");

    if($SchemaIncludes -and ($SchemaIncludes.Count -gt 0 )){
        $null = $sb.Append( " AND schema_id IN (") 
        ForEach($t in $SchemaIncludes){
          $null = $sb.Append( ("schema_id('{0}')," -f ($t -replace '[\[\]]')) )
        }
        $null = $sb.Append( " -1) `n") 
    }

    if($SchemaExcludes -and ($SchemaExcludes.Count -gt 0 )){
        $null = $sb.Append( " AND schema_id NOT IN (") 
        ForEach($t in $SchemaExcludes){
          $null = $sb.Append( ("schema_id('{0}')," -f ($t -replace '[\[\]]')) )
        }
        $null = $sb.Append( " -1) `n") 
    }

    if($TableIncludes -and ($TableIncludes.Count -gt 0 )){
        $null = $sb.Append( " AND object_id IN (") 
        ForEach($t in $TableIncludes){
          $null = $sb.Append( ("object_id('{0}')," -f ($t -replace '[\[\]]')) )
        }
        $null = $sb.Append( " -1) `n") 
    }

    if($TableExcludes -and ($TableExcludes.Count -gt 0 )){
        $null = $sb.Append( " AND object_id NOT IN (") 
        ForEach($t in $TableExcludes){
          $null = $sb.Append( ("object_id('{0}')," -f ($t -replace '[\[\]]')) )
        }
        $null = $sb.Append( " -1) `n") 
    }

    $sql = $sb.ToString()

    
    if($User){
        $results = Invoke-Sqlcmd -ServerInstance $Server -Database $Database -Query $sql -Username $User -Password $Password
    }else{
        $results = Invoke-Sqlcmd -ServerInstance $Server -Database $Database -Query $sql
    }
    ForEach($row in $results){
        $table = $row.TableName
        $tables += $table
    }

    return $tables;
}


Export-modulemember -function Invoke-BCP
Export-modulemember -function Export-Tables
Export-modulemember -function Import-Tables


