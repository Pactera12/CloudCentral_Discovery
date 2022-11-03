<#
.SYNOPSIS
Collect-ServerInfo.ps1 - PowerShell script to collect information about Windows servers

.DESCRIPTION 
This PowerShell script runs a series of WMI and other queries to collect information
about Windows servers.

.OUTPUTS
Each server's results are output to HTML.

.PARAMETER -Verbose
See more detailed progress as the script is running.

.EXAMPLE
.\Collect-ServerInfo.ps1 SERVER1
Collect information about a single server.

.EXAMPLE
"SERVER1","SERVER2","SERVER3" | .\Collect-ServerInfo.ps1
Collect information about multiple servers.

.EXAMPLE
Get-ADComputer -Filter {OperatingSystem -Like "Windows Server*"} | %{.\Collect-ServerInfo.ps1 $_.DNSHostName}
Collects information about all servers in Active Directory.


Change Log:
V1.00, 20/08/2022 - First release

#>




Begin
{
    #Initialize
     Write-Host "Initializing" start-sleep -s 5

}

Process
{

    #---------------------------------------------------------------------
    # Process each ComputerName
    #---------------------------------------------------------------------

    

     Write-Host "=====> Processing $ComputerName <====="
    
     $COMPUTERNAME = hostname


      $Date = (Get-Date).tostring("dd-MM-yyy") 

       $FolderName =  New-Item -ItemType Directory -Path C:\  -Name "$COMPUTERNAME-$Date"
       $path = 'c:\$FolderName'

    $htmlreport = @()
    $htmlbody = @()
    $htmlfile = "$($ComputerName).html"
    $spacer = "<br />"

    #---------------------------------------------------------------------
    # Do 10 pings and calculate the fastest response time
    # Not using the response time in the report yet so it might be
    # removed later.
    #---------------------------------------------------------------------
    

        #---------------------------------------------------------------------
        # Collect computer system information and convert to HTML fragment
        #---------------------------------------------------------------------
    
        Write-Host "Collecting computer system information"  start-sleep -s 5

        $subhead = "<h3>Computer System Information</h3>"
        $htmlbody += $subhead
    
        try
        {
            $csinfo = Get-WmiObject Win32_ComputerSystem -ComputerName $ComputerName -ErrorAction STOP |
                Select-Object Name,Manufacturer,Model,
                            @{Name='Physical Processors';Expression={$_.NumberOfProcessors}},
                            @{Name='Logical Processors';Expression={$_.NumberOfLogicalProcessors}},
                            @{Name='Total Physical Memory (Gb)';Expression={
                                $tpm = $_.TotalPhysicalMemory/1GB;
                                "{0:F0}" -f $tpm
                            }},
                            DnsHostName,Domain
                
            $htmlbody += $csinfo | ConvertTo-Html -Fragment
            $htmlbody += $spacer
            $csinfo | Out-File $FolderName/$computername-CSINFO.txt

            
       
        }
        catch
        {
            Write-Host $_.Exception.Message
            $htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
            $htmlbody += $spacer
        }

        $driveInfo = Get-WMIObject Win32_LogicalDisk | select name, @{n='Size';e={"{0:n2}" -f ($_.size/1gb)}},@{n='FreeSpace';e={"{0:n2}" -f ($_.freespace/1gb)}}
$driversize =   $driveinfo[0] | select size
         



        #---------------------------------------------------------------------
        # Collect operating system information and convert to HTML fragment
        #---------------------------------------------------------------------
    
        Write-Host "Collecting operating system information" start-sleep -s 5
         
        $subhead = "<h3>Operating System Information</h3>"
        $htmlbody += $subhead
    
        try
        {
            $osinfo = Get-WmiObject Win32_OperatingSystem -ComputerName $ComputerName -ErrorAction STOP | 
                Select-Object @{Name='Operating System';Expression={$_.Caption}},
                            @{Name='Architecture';Expression={$_.OSArchitecture}},
                            Version,Organization,
                            @{Name='Install Date';Expression={
                                $installdate = [datetime]::ParseExact($_.InstallDate.SubString(0,8),"yyyyMMdd",$null);
                                $installdate.ToShortDateString()
                            }},
                            WindowsDirectory

            $htmlbody += $osinfo | ConvertTo-Html -Fragment
            $htmlbody += $spacer

           $osinfo | Out-File $FolderName/$computername-OSINFO.txt
        }
        catch
        {
            Write-Warning $_.Exception.Message
            $htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
            $htmlbody += $spacer
        }


        #---------------------------------------------------------------------
        # Collect physical memory information and convert to HTML fragment
        #---------------------------------------------------------------------

        Write-Host "Collecting physical memory information"   start-sleep -s 5

        $subhead = "<h3>Physical Memory Information</h3>"
        $htmlbody += $subhead

        try
        {
            $memorybanks = @()
            $physicalmemoryinfo = @(Get-WmiObject Win32_PhysicalMemory -ComputerName $ComputerName -ErrorAction STOP |
                Select-Object DeviceLocator,Manufacturer,Speed,Capacity)

            foreach ($bank in $physicalmemoryinfo)
            {
                $memObject = New-Object PSObject
                $memObject | Add-Member NoteProperty -Name "Device Locator" -Value $bank.DeviceLocator
                $memObject | Add-Member NoteProperty -Name "Manufacturer" -Value $bank.Manufacturer
                $memObject | Add-Member NoteProperty -Name "Speed" -Value $bank.Speed
                $memObject | Add-Member NoteProperty -Name "Capacity (GB)" -Value ("{0:F0}" -f $bank.Capacity/1GB)

                $memorybanks += $memObject
            }

            $htmlbody += $memorybanks | ConvertTo-Html -Fragment
            $htmlbody += $spacer
        }
        catch
        {
            Write-Warning $_.Exception.Message
            $htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
            $htmlbody += $spacer
        }


        #---------------------------------------------------------------------
        # Collect pagefile information and convert to HTML fragment
        #---------------------------------------------------------------------

        $subhead = "<h3>PageFile Information</h3>"
        $htmlbody += $subhead

        Write-Host "Collecting pagefile information"  start-sleep -s 5

        try
        {
            $pagefileinfo = Get-WmiObject Win32_PageFileUsage -ComputerName $ComputerName -ErrorAction STOP |
                Select-Object @{Name='Pagefile Name';Expression={$_.Name}},
                            @{Name='Allocated Size (Mb)';Expression={$_.AllocatedBaseSize}}

            $htmlbody += $pagefileinfo | ConvertTo-Html -Fragment
            $htmlbody += $spacer
        }
        catch
        {
            Write-Warning $_.Exception.Message
            $htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
            $htmlbody += $spacer
        }




        #---------------------------------------------------------------------
        # Collect logical disk information and convert to HTML fragment
        #---------------------------------------------------------------------

        $subhead = "<h3>Logical Disk Information</h3>"
        $htmlbody += $subhead

        Write-Host "Collecting logical disk information"  start-sleep -s 5

        try
        {
            $diskinfo = Get-WmiObject Win32_LogicalDisk -ComputerName $ComputerName -ErrorAction STOP | 
                Select-Object DeviceID,FileSystem,VolumeName,
                @{Expression={$_.Size /1Gb -as [int]};Label="Total Size (GB)"},
                @{Expression={$_.Freespace / 1Gb -as [int]};Label="Free Space (GB)"}

            $htmlbody += $diskinfo | ConvertTo-Html -Fragment
            $htmlbody += $spacer

            $diskinfo | Out-File $FolderName/$computername-diskinfo.txt
        }
        catch
        {
            Write-Warning $_.Exception.Message
            $htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
            $htmlbody += $spacer
        }


        #---------------------------------------------------------------------
        # Collect volume information and convert to HTML fragment
        #---------------------------------------------------------------------

        $subhead = "<h3>Volume Information</h3>"
        $htmlbody += $subhead

        Write-Host "Collecting volume information"  start-sleep -s 5

        try
        {
            $volinfo = Get-WmiObject Win32_Volume -ComputerName $ComputerName -ErrorAction STOP | 
                Select-Object Label,Name,DeviceID,SystemVolume,
                @{Expression={$_.Capacity /1Gb -as [int]};Label="Total Size (GB)"},
                @{Expression={$_.Freespace / 1Gb -as [int]};Label="Free Space (GB)"}

            $htmlbody += $volinfo | ConvertTo-Html -Fragment
            $htmlbody += $spacer
            $volinfo | Out-File $FolderName/$computername-volinfo.txt
        }
        catch
        {
            Write-Warning $_.Exception.Message
            $htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
            $htmlbody += $spacer
        }


        #---------------------------------------------------------------------
        # Collect network interface information and convert to HTML fragment
        #---------------------------------------------------------------------    

        $subhead = "<h3>Network Interface Information</h3>"
        $htmlbody += $subhead

       Write-Host "Collecting network interface information"  start-sleep -s 5

        try
        {
            $nics = @()
            $nicinfo = @(Get-WmiObject Win32_NetworkAdapter -ComputerName $ComputerName -ErrorAction STOP | Where {$_.PhysicalAdapter} |
                Select-Object Name,AdapterType,MACAddress,
                @{Name='ConnectionName';Expression={$_.NetConnectionID}},
                @{Name='Enabled';Expression={$_.NetEnabled}},
                @{Name='Speed';Expression={$_.Speed/1000000}})

            $nwinfo = Get-WmiObject Win32_NetworkAdapterConfiguration -ComputerName $ComputerName -ErrorAction STOP |
                Select-Object Description, DHCPServer,  
                @{Name='IpAddress';Expression={$_.IpAddress -join '; '}},  
                @{Name='IpSubnet';Expression={$_.IpSubnet -join '; '}},  
                @{Name='DefaultIPgateway';Expression={$_.DefaultIPgateway -join '; '}},  
                @{Name='DNSServerSearchOrder';Expression={$_.DNSServerSearchOrder -join '; '}}

            foreach ($nic in $nicinfo)
            {
                $nicObject = New-Object PSObject
                $nicObject | Add-Member NoteProperty -Name "Connection Name" -Value $nic.connectionname
                $nicObject | Add-Member NoteProperty -Name "Adapter Name" -Value $nic.Name
                $nicObject | Add-Member NoteProperty -Name "Type" -Value $nic.AdapterType
                $nicObject | Add-Member NoteProperty -Name "MAC" -Value $nic.MACAddress
                $nicObject | Add-Member NoteProperty -Name "Enabled" -Value $nic.Enabled
                $nicObject | Add-Member NoteProperty -Name "Speed (Mbps)" -Value $nic.Speed
        
                $ipaddress = ($nwinfo | Where {$_.Description -eq $nic.Name}).IpAddress
                $nicObject | Add-Member NoteProperty -Name "IPAddress" -Value $ipaddress

                $nics += $nicObject
            }

            $htmlbody += $nics | ConvertTo-Html -Fragment
            $htmlbody += $spacer
        }
        catch
        {
            Write-Warning $_.Exception.Message
            $htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
            $htmlbody += $spacer
        }


        #---------------------------------------------------------------------
        # Collect software information and convert to HTML fragment
        #---------------------------------------------------------------------

        $subhead = "<h3>Software Information</h3>"
        $htmlbody += $subhead
 
        Write-Host "Collecting software information"  start-sleep -s 5
        
        try
        {
            $software = Get-WmiObject Win32_Product -ComputerName $ComputerName -ErrorAction STOP |Where-Object {$_.Vendor -notlike '*Microsoft*'} | Select-Object Vendor,Name,Version | Sort-Object Vendor,Name
        
            $htmlbody += $software | ConvertTo-Html -Fragment
            $htmlbody += $spacer 
              $software | Out-File $FolderName/$ComputerName-Non_microsoft_programs.txt
        
        }
        catch
        {
            Write-Warning $_.Exception.Message
            $htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
            $htmlbody += $spacer
        }




        $subhead = "<h3>Computer Roles Information</h3>"
        $htmlbody += $subhead
		
	Write-Host "Collecting Roles information" start-sleep -s 5

	try
	{
            $roles = Get-WindowsFeature | Where-Object {$_.InstallState -eq 'Installed' -or $_.DisplayName -Like '*Active*' -or $_.DisplayName -Like '*IIS*' } | Select-Object Name,DisplayName

            $htmlbody += $roles | ConvertTo-Html -Fragment
            $htmlbody += $spacer 
            $roles | Out-File $FolderName/$ComputerName-roles.txt
        
        }
        catch
        {
            Write-Warning $_.Exception.Message
            $htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
            $htmlbody += $spacer
        }
       
        #---------------------------------------------------------------------
        # Collect services information and covert to HTML fragment
	
        #---------------------------------------------------------------------		
		
        $subhead = "<h3>Computer Services Information</h3>"
        $htmlbody += $subhead
		 
	Write-Host "Collecting services information" start-sleep -s 5

	try
	{
            $services = Get-WmiObject Win32_Service -ComputerName $ComputerName -ErrorAction STOP  |Where {$_.DisplayName -Like "*Exchange*" -or $_.DisplayName -like "MSSQL$*" -or $_.DisplayName -like "MSSQLSERVER" -or $_.Name -Like "*sql*" } | Select-Object Name,StartName,State,StartMode | Sort-Object Name

            $htmlbody += $services | ConvertTo-Html -Fragment
            $htmlbody += $spacer 
            $services | Out-File $FolderName/$ComputerName-services.txt
        
        }
        catch
        {
            Write-Warning $_.Exception.Message
            $htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
            $htmlbody += $spacer
        }

        #---------------------------------------------------------------------

     try
	{
            $iis = Get-IISSITE

            $htmlbody += $iis | ConvertTo-Html -Fragment
            $htmlbody += $spacer 
            $iis | Out-File $FolderName/$ComputerName-iis.txt
        
        }
        catch
        {
            Write-Warning $_.Exception.Message
            $htmlbody += "<p>An error was encountered. $($_.Exception.Message)</p>"
            $htmlbody += $spacer
        }

        # Generate the HTML report and output to file
        #---------------------------------------------------------------------
	
        Write-Host "Producing HTML report" start-sleep -s 5
    
        $reportime = Get-Date

        #Common HTML head and styles
	    $htmlhead="<html>
				    <style>
				    BODY{font-family: Arial; font-size: 8pt;}
				    H1{font-size: 20px;}
				    H2{font-size: 18px;}
				    H3{font-size: 16px;}
				    TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
				    TH{border: 1px solid black; background: #dddddd; padding: 5px; color: #000000;}
				    TD{border: 1px solid black; padding: 5px; }
				    td.pass{background: #7FFF00;}
				    td.warn{background: #FFE600;}
				    td.fail{background: #FF0000; color: #ffffff;}
				    td.info{background: #85D4FF;}
				    </style>
				    <body>
				    <h1 align=""center"">Server Info: $ComputerName</h1>
				    <h3 align=""center"">Generated: $reportime</h3>"

        $htmltail = "</body>
			    </html>"

        $htmlreport = $htmlhead + $htmlbody + $htmltail

        $htmlreport | Out-File $htmlfile -Encoding Utf8





              $Body = @{
    # id = "1"
     os =  $osinfo.'Operating System'
     Domain = $csinfo.domain
     deviceId = "1"
     systemInfo = $csinfo.Name
     diskTotalSize = $driversize.size
     logicalProcessor = $csinfo.'Logical Processors' 
     TotalPhysicalMemory = $csinfo.'Total Physical Memory (Gb)'

 }
 $JsonBody = $Body | ConvertTo-Json



 $Params = @{
     Method = "PUT"
     Uri = "https://retoolapi.dev/n06gcX/serverInfo/1"
     Body = $JsonBody
     ContentType = "application/json"
 }
 Invoke-RestMethod @Params
                
    


    
 $RAM = (Get-CimInstance Win32_PhysicalMemory | Measure-Object -Property capacity -Sum).sum /1gb

  $roles = Get-WindowsFeature | Where-Object {($_.InstallState -eq 'Installed') -and (($_.DisplayName -Like '*Active*')  )} 
  $hostname = hostname

 $service =  Get-Service | Where {$_.DisplayName -Like "*Exchange*" -or $_.DisplayName -like "MSSQL$*" -or $_.DisplayName -like "MSSQLSERVER" -or $_.Name -Like "*sql*"  -or $_.Name -Like "*IIS*"} 

 if($RAM  -lt 9 -and -NOT $service -and -Not  $roles )  {

Write-Output "{$hostname   :   Simple }"  | Out-File  test.txt
 
 }elseif(($RAM -gt 9 -and $RAM -lt 15) -or  $service -and -Not $roles) {
 
  Write-Output "{$hostname   :   Medium }" | Out-File  test.txt
 }

 elseif(($RAM -gt 16) -and  $service -or  $roles) {
 
  Write-Output "{$hostname   :   Complex }" | Out-File  test.txt
 }else{
  Write-Output "This is Medium" | Out-File  test.txt
 }





End
{
    #Wrap it up
    Write-Host "=====> Finished <=====" 
}



}