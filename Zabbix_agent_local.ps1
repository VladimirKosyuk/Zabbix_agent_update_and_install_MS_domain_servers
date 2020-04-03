#Created by https://github.com/VladimirKosyuk

#Zabbix agent installation and update for Microsoft domain servers
#
# Build date: 04.03.2020									   
 
#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Scenarios:
<# 
no zabbix installed
no firewall rule is set
no binares, service exist
no service, binares exist
service exist, binares are broken
binares exist, service is broken path
binares and service are broken
no binares, service are broken
local repo is empty, service exist
no service, binares are broken
service ok, config is broken
zabbix agent version is not equal to repo version
zabbix conf is not equal to repo version, Hostname string is excluded from compare
#>

#Codnitions:
<# 
Module Active directory;
Windows 2012 R2 (for Windows 2008 R2 cannot create firewall rule and check connectivity between host and zabbix server);
Default path for zabbix agent need to be c:\Program Files\zabbix40
Repository need to placed in NETLOGON\zabbix40
#>

#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

$confg = "$PSScriptRoot\Zbx_agent"+('.txt')
$globallog = "$PSScriptRoot\Zbx_agent"+('.log')

try

{
    $values = (Get-Content $confg).Replace( '\', '\\') | ConvertFrom-StringData 
    $Output = $values.Output
  
}

catch

{   
    Write-Host "No config file has been found" -ForegroundColor RED
    Write-Output $Error[0].Exception.Message
    Write-Host "Script terminated with error." -ForegroundColor Yellow
    Write-Output ("FATAL ERROR - Config file is accessible check not passed  "+(Get-Date)) | Out-File "$globallog" -Append
    Break 
}


#vars

$Debag_Date = Get-Date -Format "MM.dd.yyyy"
$Domain_name = Get-WmiObject -Class Win32_ComputerSystem |Select-Object -ExpandProperty "Domain"
$Local_name = (($env:computername).ToLower())+"."+(($Domain_name).ToLower())
$Global_Repo = "\\"+$Domain_name+"\NETLOGON\zabbix40"
$Local_Repo = "c:\Program Files\zabbix40"

#1 find zabbix server
if ($ZBX_service = Get-WmiObject win32_service | Where-Object{($_.PathName -like '*zabbix*')}) 
#Service must be running
    {If ($ZBX_service.State -match "Running") 
    #Check firewall, connectivity, check exe version, check if active config string is set
    {
        #Check rule for passive checks, if not exist - create
        if (!(Get-NetFirewallPortFilter -PolicyStore ActiveStore | where {$_.LocalPort -match "10050"})) 
            {

            try

            {
            New-NetFirewallRule -DisplayName 'Zabbix Agent' -Profile @('Domain', ,'Public','Private') -Direction Inbound -Action Allow -Protocol TCP -LocalPort @('10050')
            Write-Output ("INFORMATION"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+"Firewall rule is set for passive checks"+";"+"id_1") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
            }

            catch

            {
            Write-Output ("WARNING"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($Error[0].Exception.Message )+";"+"id_2") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
            }
        }
        #Check rule for active checks, if not exist - create
        if (!(Get-NetFirewallPortFilter -PolicyStore ActiveStore | where {$_.LocalPort -match "10051"})) 
            {

            try

            {
            New-NetFirewallRule -DisplayName 'Zabbix Agent' -Profile @('Domain', ,'Public','Private') -Direction Outbound -Action Allow -Protocol TCP -LocalPort @('10051')
            Write-Output ("OK"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+"Firewall rule is set for active checks"+";"+"id_3") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
            }

            catch

            {
            Write-Output ("WARNING"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($Error[0].Exception.Message )+";"+"id_4") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
            }
        }
        #Check connectivity
        $ServerActive = (((Get-Content ($Local_Repo+"\ZABBIX_AGENTD.WIN.CONF") | select-string -pattern "^ServerActive=" -casesensitive |Select-Object -ExpandProperty line ) -replace ‘[ServerActive=]’,”").split(",")) -replace '\s',''
        foreach ($ServerActive_Unit in $ServerActive) 
            {
            if 
            ((Test-NetConnection -ComputerName $ServerActive_Unit -Port '10051' -WarningVariable bad_active| select ComputerName, TcpTestSucceeded)) 
                {if 
                    ($bad_active) 
                        {Write-Output ("WARNING"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($bad_active -join ",")+";"+"id_5") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append}
                }
}
        #check exe version, if not equal to repository version - update
        if (diff(get-childitem ($Global_Repo+"\ZABBIX_AGENTD.EXE") | Select-Object -ExpandProperty VersionInfo | Select-Object -ExpandProperty ProductVersion)(get-childitem ($Local_Repo+"\ZABBIX_AGENTD.EXE") | Select-Object -ExpandProperty VersionInfo | Select-Object -ExpandProperty ProductVersion))
            {
            Write-Output ("INFORMATION"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+"Updating ZABBIX_AGENTD.EXE - version is differ to"+($Global_Repo)+";"+"id_6") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
            Stop-Service -name (Get-WmiObject win32_service | Where-Object{($_.PathName -like '*zabbix*')}| Select-Object -ExpandProperty "Name") -ErrorAction SilentlyContinue

                try 

                            {
                            Get-Process | Where-Object{($_.Name -like '*zabbix*')} | Stop-Process -Force
                            Remove-Item -Path ($Local_Repo+"\ZABBIX_AGENTD.EXE") -Force | out-null
                            }

                        catch 

                            {
                            Write-Output ("Error"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($Error[0].Exception.Message )+";"+"id_7") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
                            Break
                            }

                    #copy from global repo
                    #here must be try\catch and break if fails
                    If (Test-Path $Global_Repo)
                        {

                        try

                            {
                            Copy-Item -Path ($Global_Repo+"\ZABBIX_AGENTD.EXE") -Destination $Local_Repo -Recurse
                            }

                        catch

                            {
                            Write-Output ("Error"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($Error[0].Exception.Message )+";"+"id_8") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
                            Break
                            }
                        }

                    else

                        {
                        Write-Output ("Error"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+"Cannot access to "+($Global_Repo)+";"+"id_9") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
                        Break
                        }
                    #output actions to the log
                    Write-Output ("INFORMATION"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($Local_Repo)+" ZABBIX_AGENTD.EXE has been updated from "+($Global_Repo)+";"+"id_10") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
                    #try register and start service
                    try

                        {
                        Start-Service -name (Get-WmiObject win32_service | Where-Object{($_.PathName -like '*zabbix*')}| Select-Object -ExpandProperty "Name")
                        #zabbix agent has been registered and running
                        Write-Output ("OK"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($ZBX_service.Name)+" service has been updated"+";"+"id_11") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append 
                        }

                    catch

                        {
                        #Cannot register and run zabbix_agentd.exe as service 
                        Write-Output ("Error"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($Error[0].Exception.Message )+";"+"id_12") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
                        Break
                        }
            
    }
        #check FQDN in ZABBIX_AGENTD.WIN.CONF, must be equal to real FQDN
        if ($Hostname = Get-Content ($Local_Repo+"\ZABBIX_AGENTD.WIN.CONF") | select-string -pattern "^Hostname=" -casesensitive |Select-Object -ExpandProperty line)
            {
            if (!(diff($Hostname)($Local_name)))
                {
                   try

                    {(Get-Content ($Local_Repo+"\ZABBIX_AGENTD.WIN.CONF")) -replace $Hostname, ("Hostname="+$Local_name) | Set-Content ($Local_Repo+"\ZABBIX_AGENTD.WIN.CONF")
                    Restart-Service -name (Get-WmiObject win32_service | Where-Object{($_.PathName -like '*zabbix*')}| Select-Object -ExpandProperty "Name")
                    Write-Output ("OK"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($ZBX_service.Name)+" FQDN has been set to ZABBIX_AGENTD.WIN.CONF"+";"+"id_13") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
                    }
    
                   catch

                    {
                    Write-Output ("WARNING"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($Error[0].Exception.Message )+";"+"id_14") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
                    }
                }

            }
        #If no FQDN defined - need to add
        else 
            {
                try

                    {
                    Add-Content -Value ("Hostname="+$Local_name) -Path ($Local_Repo+"\ZABBIX_AGENTD.WIN.CONF")
                    Restart-Service -name (Get-WmiObject win32_service | Where-Object{($_.PathName -like '*zabbix*')}| Select-Object -ExpandProperty "Name")
                    Write-Output ("OK"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($ZBX_service.Name)+" FQDN has been added to ZABBIX_AGENTD.WIN.CONF"+";"+"id_15") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
                    }
    
                catch

                    {Write-Output ("WARNING"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($Error[0].Exception.Message )+";"+"id_16") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append}
            }
        #check ZABBIX_AGENTD.WIN.CONF
        if (diff(get-content ($Global_Repo+"\ZABBIX_AGENTD.WIN.CONF"))(get-Content ($Local_Repo+"\ZABBIX_AGENTD.WIN.CONF") | select-string -notmatch -pattern "^Hostname=" -casesensitive))
            {
    Write-Output ("INFORMATION"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+"Updating ZABBIX_AGENTD.WIN.CONF - version is differ to"+($Global_Repo)+";"+"id_X1") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
        try

            {
            Set-Content -Value (get-content ($Global_Repo+"\ZABBIX_AGENTD.WIN.CONF")) -Path ($Local_Repo+"\ZABBIX_AGENTD.WIN.CONF")
            Add-Content -Value ("Hostname="+$Local_name) -Path ($Local_Repo+"\ZABBIX_AGENTD.WIN.CONF")
            Restart-Service -name (Get-WmiObject win32_service | Where-Object{($_.PathName -like '*zabbix*')}| Select-Object -ExpandProperty "Name")
            Write-Output ("OK"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+"ZABBIX_AGENTD.WIN.CONF is updated"+";"+"id_X2") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
            }
    
        catch

            {
            Write-Output ("Error"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($Error[0].Exception.Message )+";"+"id_X3") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
            Break
            }

    }
    }
    
    else

    #3 Если найдена, но статус не running - запустить
    {
        #service start
        try

            {
            #Zabbix Agent service service has been found but not running, trying to start service
            Start-Service -name $ZBX_service.name -ErrorAction Stop
            Write-Output ("Information"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($ZBX_service.Name)+" service has been pushed to run"+";"+"id_17") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append 
            }
        #cannot start service - unregister, delete local repo, copy repo from global, register and start service
        catch

            { 
            #Cannot start Zabbix Agent service
            Write-Output ("WARNING"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($Error[0].Exception.Message)+";"+"id_18") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append 
            #stop services that may block service deletion and delete service
            Get-Process | Where-Object{(($_.Name -like '*taskmgr*') -or ($_.Name -like '*zabbix*') -or ($_.Name -like '*mmc*'))} | Stop-Process -Force
            $ZBX_service.delete()
            #delete local binares, if exist
            If (Test-Path $Local_Repo) 
                {
                #here must be try\catch and break if fails
                try 

                    {
                    Get-Process | Where-Object{($_.Name -like '*zabbix*')} | Stop-Process -Force
                    Remove-Item -Path $Local_Repo -recurse -Force | out-null
                    }

                catch 

                    {
                    Write-Output ("Error"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($Error[0].Exception.Message )+";"+"id_19") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
                    Break
                    }
                }

            else

                {
                Write-Output ("INFORMATION"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+"Local binares does not exist, trying to create"+";"+"id_20") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
                }
            #create local repo folder
            New-Item -ItemType directory -Path $Local_Repo | out-null
            #copy from global repo
            #here must be try\catch and break if fails
            If (Test-Path $Global_Repo)
                {

                try

                    {
                    Copy-Item -Path $Global_Repo"\*" -Destination $Local_Repo -Recurse
                    }

                catch

                    {
                    Write-Output ("Error"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($Error[0].Exception.Message )+";"+"id_21") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
                    Break
                    }
                }

            else

                {
                Write-Output ("Error"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+"Cannot access to "+($Global_Repo)+";"+"id_22") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
                Break
                }
            #output actions to the log
            Write-Output ("INFORMATION"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($Local_Repo)+" has been recreated and updated from "+($Global_Repo)+";"+"id_23") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
            #try register and start service
            try

                {
                New-Service -Name "Zabbix Agent" -BinaryPathName '"c:\Program Files\zabbix40\zabbix_agentd.exe" --config "c:\Program Files\zabbix40\zabbix_agentd.win.conf"'  -DisplayName "Zabbix Agent" -StartupType Automatic -Description "Provides system monitoring via zabbix" -ErrorAction Stop| Out-Null
                Start-Service -name "Zabbix Agent" -ErrorAction Stop
                #zabbix agent has been registered and running
                Write-Output ("OK"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($ZBX_service.Name)+" service has been registered and pushed to run"+";"+"id_24") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append 
                }

            catch

                {
                #Cannot register and run zabbix_agentd.exe as service 
                Write-Output ("Error"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($Error[0].Exception.Message )+";"+"id_25") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
                Break
                }
            }
        #Check rule for passive checks, if not exist - create
        if (!(Get-NetFirewallPortFilter -PolicyStore ActiveStore | where {$_.LocalPort -match "10050"})) 
            {

            try

            {
            New-NetFirewallRule -DisplayName 'Zabbix Agent' -Profile @('Domain', ,'Public','Private') -Direction Inbound -Action Allow -Protocol TCP -LocalPort @('10050')
            Write-Output ("OK"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+"Firewall rule is set for passive checks"+";"+"id_26") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
            }

            catch

            {
            Write-Output ("WARNING"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($Error[0].Exception.Message )+";"+"id_27") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
            }
        }
        #Check rule for active checks, if not exist - create
        if (!(Get-NetFirewallPortFilter -PolicyStore ActiveStore | where {$_.LocalPort -match "10051"})) 
            {

            try

            {
            New-NetFirewallRule -DisplayName 'Zabbix Agent' -Profile @('Domain', ,'Public','Private') -Direction Outbound -Action Allow -Protocol TCP -LocalPort @('10051')
            Write-Output ("OK"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+"Firewall rule is set for active checks"+";"+"id_28") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
            }

            catch

            {
            Write-Output ("WARNING"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($Error[0].Exception.Message )+";"+"id_29") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
            }
        }
        #Check connectivity
        $ServerActive = (((Get-Content ($Local_Repo+"\ZABBIX_AGENTD.WIN.CONF") | select-string -pattern "^ServerActive=" -casesensitive |Select-Object -ExpandProperty line ) -replace ‘[ServerActive=]’,”").split(",")) -replace '\s',''
        foreach ($ServerActive_Unit in $ServerActive) 
            {
            if 
            ((Test-NetConnection -ComputerName $ServerActive_Unit -Port '10051' -WarningVariable bad_active| select ComputerName, TcpTestSucceeded)) 
                {if 
                    ($bad_active) 
                        {Write-Output ("WARNING"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($bad_active -join ",")+";"+"id_30") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append}
                }
}
        #check exe version, if not equal to repository version - update
        if (diff(get-childitem ($Global_Repo+"\ZABBIX_AGENTD.EXE") | Select-Object -ExpandProperty VersionInfo | Select-Object -ExpandProperty ProductVersion)(get-childitem ($Local_Repo+"\ZABBIX_AGENTD.EXE") | Select-Object -ExpandProperty VersionInfo | Select-Object -ExpandProperty ProductVersion))
            {
            Write-Output ("INFORMATION"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+"Updating ZABBIX_AGENTD.EXE - version is differ to"+($Global_Repo)+";"+"id_31") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
            Stop-Service -name (Get-WmiObject win32_service | Where-Object{($_.PathName -like '*zabbix*')}| Select-Object -ExpandProperty "Name") -ErrorAction SilentlyContinue

                try 

                            {
                            Get-Process | Where-Object{($_.Name -like '*zabbix*')} | Stop-Process -Force
                            Remove-Item -Path ($Local_Repo+"\ZABBIX_AGENTD.EXE") -Force | out-null
                            }

                        catch 

                            {
                            Write-Output ("Error"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($Error[0].Exception.Message )+";"+"id_32") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
                            Break
                            }

                    #copy from global repo
                    #here must be try\catch and break if fails
                    If (Test-Path $Global_Repo)
                        {

                        try

                            {
                            Copy-Item -Path ($Global_Repo+"\ZABBIX_AGENTD.EXE") -Destination $Local_Repo -Recurse
                            }

                        catch

                            {
                            Write-Output ("Error"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($Error[0].Exception.Message )+";"+"id_33") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
                            Break
                            }
                        }

                    else

                        {
                        Write-Output ("Error"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+"Cannot access to "+($Global_Repo)+";"+"id_34") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
                        Break
                        }
                    #output actions to the log
                    Write-Output ("INFORMATION"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($Local_Repo)+" ZABBIX_AGENTD.EXE has been updated from "+($Global_Repo)+";"+"id_35") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
                    #try register and start service
                    try

                        {
                        Start-Service -name (Get-WmiObject win32_service | Where-Object{($_.PathName -like '*zabbix*')}| Select-Object -ExpandProperty "Name")
                        #zabbix agent has been registered and running
                        Write-Output ("OK"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($ZBX_service.Name)+" service has been updated"+";"+"id_36") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append 
                        }

                    catch

                        {
                        #Cannot register and run zabbix_agentd.exe as service 
                        Write-Output ("Error"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($Error[0].Exception.Message )+";"+"id_37") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
                        Break
                        }
            
    }
        #check FQDN in ZABBIX_AGENTD.WIN.CONF, must be equal to real FQDN
        if ($Hostname = Get-Content ($Local_Repo+"\ZABBIX_AGENTD.WIN.CONF") | select-string -pattern "^Hostname=" -casesensitive |Select-Object -ExpandProperty line)
            {
            if (!(diff($Hostname)($Local_name)))
                {
                   try

                    {(Get-Content ($Local_Repo+"\ZABBIX_AGENTD.WIN.CONF")) -replace $Hostname, ("Hostname="+$Local_name) | Set-Content ($Local_Repo+"\ZABBIX_AGENTD.WIN.CONF")
                    Restart-Service -name (Get-WmiObject win32_service | Where-Object{($_.PathName -like '*zabbix*')}| Select-Object -ExpandProperty "Name")
                    Write-Output ("OK"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($ZBX_service.Name)+" FQDN has been set to ZABBIX_AGENTD.WIN.CONF"+";"+"id_38") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
                    }
    
                   catch

                    {
                    Write-Output ("WARNING"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($Error[0].Exception.Message )+";"+"id_39") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
                    }
                }

            }
        #If no FQDN defined - need to add
        else 
            {
                try

                    {
                    Add-Content -Value ("Hostname="+$Local_name) -Path ($Local_Repo+"\ZABBIX_AGENTD.WIN.CONF")
                    Restart-Service -name (Get-WmiObject win32_service | Where-Object{($_.PathName -like '*zabbix*')}| Select-Object -ExpandProperty "Name")
                    Write-Output ("OK"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($ZBX_service.Name)+" FQDN has been added to ZABBIX_AGENTD.WIN.CONF"+";"+"id_40") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
                    }
    
                catch

                    {Write-Output ("WARNING"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($Error[0].Exception.Message )+";"+"id_41") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append}
            }
        #check ZABBIX_AGENTD.WIN.CONF
        if (diff(get-content ($Global_Repo+"\ZABBIX_AGENTD.WIN.CONF"))(get-Content ($Local_Repo+"\ZABBIX_AGENTD.WIN.CONF") | select-string -notmatch -pattern "^Hostname=" -casesensitive))
            {
    Write-Output ("INFORMATION"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+"Updating ZABBIX_AGENTD.WIN.CONF - version is differ to"+($Global_Repo)+";"+"id_X4") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
        try

            {
            Set-Content -Value (get-content ($Global_Repo+"\ZABBIX_AGENTD.WIN.CONF")) -Path ($Local_Repo+"\ZABBIX_AGENTD.WIN.CONF")
            Add-Content -Value ("Hostname="+$Local_name) -Path ($Local_Repo+"\ZABBIX_AGENTD.WIN.CONF")
            Restart-Service -name (Get-WmiObject win32_service | Where-Object{($_.PathName -like '*zabbix*')}| Select-Object -ExpandProperty "Name")
            Write-Output ("OK"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+"ZABBIX_AGENTD.WIN.CONF is updated"+";"+"id_X5") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
            }
    
        catch

            {
            Write-Output ("Error"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($Error[0].Exception.Message )+";"+"id_X6") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
            Break
            }

    }

    }

    }
#Service not found - check local repo, if exist - register service,if cannot - delete local repo, copy repo from global repo, register service
else 
    {
    #output actions to the log
    Write-Output ("INFORMATION"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+"Zabbix Agent service not found, trying to create"+";"+"id_42") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
    #if local repo exists, need to try register service
    If (Test-Path $Local_Repo)
        {
        #try to start service
        try
        
            {
            New-Service -Name "Zabbix Agent" -BinaryPathName '"c:\Program Files\zabbix40\zabbix_agentd.exe" --config "c:\Program Files\zabbix40\zabbix_agentd.win.conf"'  -DisplayName "Zabbix Agent" -StartupType Automatic -Description "Provides system monitoring via zabbix" -ErrorAction Stop| Out-Null
            Start-Service -name "Zabbix Agent" -ErrorAction Stop
            #zabbix agent has been registered and running
            Write-Output ("OK"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($ZBX_service.Name)+" service has been registered and pushed to run"+";"+"id_43") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append 
            }
        #cannot start - delete local repo, copy repo from global repo, register service
        catch

            {
            #Cannot register and run zabbix_agentd.exe as service 
            Write-Output ("Warning"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($Error[0].Exception.Message )+";"+"id_44") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append 
            #delete local binares
            try 

                    {
                    Get-Process | Where-Object{($_.Name -like '*zabbix*')} | Stop-Process -Force
                    Remove-Item -Path $Local_Repo -recurse -Force | out-null
                    }
                    
            catch 

                    {
                    Write-Output ("Error"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($Error[0].Exception.Message )+";"+"id_45") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
                    Break
                    }
            #create local repo folder
            New-Item -ItemType directory -Path $Local_Repo | out-null
            #copy from global repo
            If (Test-Path $Global_Repo)
                {

                try

                    {
                    Copy-Item -Path $Global_Repo"\*" -Destination $Local_Repo -Recurse
                    }

                catch

                    {
                    Write-Output ("Error"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($Error[0].Exception.Message )+";"+"id_46") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
                    Break
                    }
                }

            else

                {
                Write-Output ("Error"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+"Cannot access to "+($Global_Repo)+";"+"id_47") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
                Break
                }
            #output actions to the log
            Write-Output ("INFORMATION"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($Local_Repo)+" has been recreated and updated from "+($Global_Repo)+";"+"id_48") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
            #try  start service
            try

                {
                Start-Service -name "Zabbix Agent" -ErrorAction Stop
                #zabbix agent has been registered and running
                Write-Output ("OK"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($ZBX_service.Name)+" service has been registered and pushed to run"+";"+"id_49") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append 
                }

            catch

                {
                #Cannot register and run zabbix_agentd.exe as service 
                Write-Output ("Error"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($Error[0].Exception.Message )+";"+"id_50") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
                Break
                }
            }
        #Check rule for passive checks, if not exist - create
        if (!(Get-NetFirewallPortFilter -PolicyStore ActiveStore | where {$_.LocalPort -match "10050"})) 
            {

            try

            {
            New-NetFirewallRule -DisplayName 'Zabbix Agent' -Profile @('Domain', ,'Public','Private') -Direction Inbound -Action Allow -Protocol TCP -LocalPort @('10050')
            Write-Output ("OK"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+"Firewall rule is set for passive checks"+";"+"id_51") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
            }

            catch

            {
            Write-Output ("WARNING"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($Error[0].Exception.Message )+";"+"id_52") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
            }
        }
        #Check rule for active checks, if not exist - create
        if (!(Get-NetFirewallPortFilter -PolicyStore ActiveStore | where {$_.LocalPort -match "10051"})) 
            {

            try

            {
            New-NetFirewallRule -DisplayName 'Zabbix Agent' -Profile @('Domain', ,'Public','Private') -Direction Outbound -Action Allow -Protocol TCP -LocalPort @('10051')
            Write-Output ("OK"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+"Firewall rule is set for active checks"+";"+"id_53") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
            }

            catch

            {
            Write-Output ("WARNING"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($Error[0].Exception.Message )+";"+"id_54") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
            }
        }
        #Check connectivity
        $ServerActive = (((Get-Content ($Local_Repo+"\ZABBIX_AGENTD.WIN.CONF") | select-string -pattern "^ServerActive=" -casesensitive |Select-Object -ExpandProperty line ) -replace ‘[ServerActive=]’,”").split(",")) -replace '\s',''
        foreach ($ServerActive_Unit in $ServerActive) 
            {
            if 
            ((Test-NetConnection -ComputerName $ServerActive_Unit -Port '10051' -WarningVariable bad_active| select ComputerName, TcpTestSucceeded)) 
                {if 
                    ($bad_active) 
                        {Write-Output ("WARNING"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($bad_active -join ",")+";"+"id_55") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append}
                }
}
        #check exe version, if not equal to repository version - update
        if (diff(get-childitem ($Global_Repo+"\ZABBIX_AGENTD.EXE") | Select-Object -ExpandProperty VersionInfo | Select-Object -ExpandProperty ProductVersion)(get-childitem ($Local_Repo+"\ZABBIX_AGENTD.EXE") | Select-Object -ExpandProperty VersionInfo | Select-Object -ExpandProperty ProductVersion))
            {
            Write-Output ("INFORMATION"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+"Updating ZABBIX_AGENTD.EXE - version is differ to"+($Global_Repo)+";"+"id_56") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
            Stop-Service -name (Get-WmiObject win32_service | Where-Object{($_.PathName -like '*zabbix*')}| Select-Object -ExpandProperty "Name") -ErrorAction SilentlyContinue

                try 

                            {
                            Get-Process | Where-Object{($_.Name -like '*zabbix*')} | Stop-Process -Force
                            Remove-Item -Path ($Local_Repo+"\ZABBIX_AGENTD.EXE") -Force | out-null
                            }

                        catch 

                            {
                            Write-Output ("Error"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($Error[0].Exception.Message )+";"+"id_57") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
                            Break
                            }

                    #copy from global repo
                    #here must be try\catch and break if fails
                    If (Test-Path $Global_Repo)
                        {

                        try

                            {
                            Copy-Item -Path ($Global_Repo+"\ZABBIX_AGENTD.EXE") -Destination $Local_Repo -Recurse
                            }

                        catch

                            {
                            Write-Output ("Error"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($Error[0].Exception.Message )+";"+"id_58") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
                            Break
                            }
                        }

                    else

                        {
                        Write-Output ("Error"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+"Cannot access to "+($Global_Repo)+";"+"id_59") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
                        Break
                        }
                    #output actions to the log
                    Write-Output ("INFORMATION"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($Local_Repo)+" ZABBIX_AGENTD.EXE has been updated from "+($Global_Repo)+";"+"id_60") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
                    #try register and start service
                    try

                        {
                        Start-Service -name (Get-WmiObject win32_service | Where-Object{($_.PathName -like '*zabbix*')}| Select-Object -ExpandProperty "Name")
                        #zabbix agent has been registered and running
                        Write-Output ("OK"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($ZBX_service.Name)+" service has been updated"+";"+"id_61") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append 
                        }

                    catch

                        {
                        #Cannot register and run zabbix_agentd.exe as service 
                        Write-Output ("Error"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($Error[0].Exception.Message )+";"+"id_62") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
                        Break
                        }
            
    }
        #check FQDN in ZABBIX_AGENTD.WIN.CONF, must be equal to real FQDN
        if ($Hostname = Get-Content ($Local_Repo+"\ZABBIX_AGENTD.WIN.CONF") | select-string -pattern "^Hostname=" -casesensitive |Select-Object -ExpandProperty line)
            {
            if (!(diff($Hostname)($Local_name)))
                {
                   try

                    {(Get-Content ($Local_Repo+"\ZABBIX_AGENTD.WIN.CONF")) -replace $Hostname, ("Hostname="+$Local_name) | Set-Content ($Local_Repo+"\ZABBIX_AGENTD.WIN.CONF")
                    Restart-Service -name (Get-WmiObject win32_service | Where-Object{($_.PathName -like '*zabbix*')}| Select-Object -ExpandProperty "Name")
                    Write-Output ("OK"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($ZBX_service.Name)+" FQDN has been set to ZABBIX_AGENTD.WIN.CONF"+";"+"id_63") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
                    }
    
                   catch

                    {
                    Write-Output ("WARNING"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($Error[0].Exception.Message )+";"+"id_64") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
                    }
                }

            }
        #If no FQDN defined - need to add
        else 
            {
                try

                    {
                    Add-Content -Value ("Hostname="+$Local_name) -Path ($Local_Repo+"\ZABBIX_AGENTD.WIN.CONF")
                    Restart-Service -name (Get-WmiObject win32_service | Where-Object{($_.PathName -like '*zabbix*')}| Select-Object -ExpandProperty "Name")
                    Write-Output ("OK"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($ZBX_service.Name)+" FQDN has been added to ZABBIX_AGENTD.WIN.CONF"+";"+"id_65") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
                    }
    
                catch

                    {Write-Output ("WARNING"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($Error[0].Exception.Message )+";"+"id_66") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append}
            }
        #check ZABBIX_AGENTD.WIN.CONF
        if (diff(get-content ($Global_Repo+"\ZABBIX_AGENTD.WIN.CONF"))(get-Content ($Local_Repo+"\ZABBIX_AGENTD.WIN.CONF") | select-string -notmatch -pattern "^Hostname=" -casesensitive))
            {
    Write-Output ("INFORMATION"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+"Updating ZABBIX_AGENTD.WIN.CONF - version is differ to"+($Global_Repo)+";"+"id_X7") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
        try

            {
            Set-Content -Value (get-content ($Global_Repo+"\ZABBIX_AGENTD.WIN.CONF")) -Path ($Local_Repo+"\ZABBIX_AGENTD.WIN.CONF")
            Add-Content -Value ("Hostname="+$Local_name) -Path ($Local_Repo+"\ZABBIX_AGENTD.WIN.CONF")
            Restart-Service -name (Get-WmiObject win32_service | Where-Object{($_.PathName -like '*zabbix*')}| Select-Object -ExpandProperty "Name")
            Write-Output ("OK"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+"ZABBIX_AGENTD.WIN.CONF is updated"+";"+"id_X8") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
            }
    
        catch

            {
            Write-Output ("Error"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($Error[0].Exception.Message )+";"+"id_X9") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
            Break
            }

    }

        }
    #if local repo not exist - copy from global repo and register service
    else
        {
        #output actions to the log
        Write-Output ("INFORMATION"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+"Local binares does not exist, trying to create"+";"+"id_67") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
        New-Item -ItemType directory -Path $Local_Repo | out-null
            #copy from global repo
            If (Test-Path $Global_Repo)
                {

                try

                    {
                    Copy-Item -Path $Global_Repo"\*" -Destination $Local_Repo -Recurse
                    }

                catch

                    {
                    Write-Output ("Error"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($Error[0].Exception.Message )+";"+"id_68") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
                    Break
                    }
                }

            else

                {
                Write-Output ("Error"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+"Cannot access to "+($Global_Repo)+";"+"id_69") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
                Break
                }
            #output actions to the log
            Write-Output ("INFORMATION"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($Local_Repo)+" has been recreated and updated from "+($Global_Repo)+";"+"id_70") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
            #try register and start service
            try

                {
                New-Service -Name "Zabbix Agent" -BinaryPathName '"c:\Program Files\zabbix40\zabbix_agentd.exe" --config "c:\Program Files\zabbix40\zabbix_agentd.win.conf"'  -DisplayName "Zabbix Agent" -StartupType Automatic -Description "Provides system monitoring via zabbix" -ErrorAction Stop| Out-Null
                Start-Service -name "Zabbix Agent" -ErrorAction Stop
                #zabbix agent has been registered and running
                Write-Output ("OK"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($ZBX_service.Name)+" service has been registered and pushed to run"+";"+"id_71") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append 
                }

            catch

                {
                #Cannot register and run zabbix_agentd.exe as service 
                Write-Output ("Error"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($Error[0].Exception.Message )+";"+"id_72") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
                Break
                }
            #Check rule for passive checks, if not exist - create
            if (!(Get-NetFirewallPortFilter -PolicyStore ActiveStore | where {$_.LocalPort -match "10050"})) 
                {

            try

            {
            New-NetFirewallRule -DisplayName 'Zabbix Agent' -Profile @('Domain', ,'Public','Private') -Direction Inbound -Action Allow -Protocol TCP -LocalPort @('10050')
            Write-Output ("OK"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+"Firewall rule is set for passive checks"+";"+"id_73") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
            }

            catch

            {
            Write-Output ("WARNING"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($Error[0].Exception.Message )+";"+"id_74") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
            }
        }
            #Check rule for active checks, if not exist - create
            if (!(Get-NetFirewallPortFilter -PolicyStore ActiveStore | where {$_.LocalPort -match "10051"})) 
                {

            try

            {
            New-NetFirewallRule -DisplayName 'Zabbix Agent' -Profile @('Domain', ,'Public','Private') -Direction Outbound -Action Allow -Protocol TCP -LocalPort @('10051')
            Write-Output ("OK"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+"Firewall rule is set for active checks"+";"+"id_75") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
            }

            catch

            {
            Write-Output ("WARNING"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($Error[0].Exception.Message )+";"+"id_76") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
            }
        }
            #Check connectivity
            $ServerActive = (((Get-Content ($Local_Repo+"\ZABBIX_AGENTD.WIN.CONF") | select-string -pattern "^ServerActive=" -casesensitive |Select-Object -ExpandProperty line ) -replace ‘[ServerActive=]’,”").split(",")) -replace '\s',''
            foreach ($ServerActive_Unit in $ServerActive) 
                {
                if 
                ((Test-NetConnection -ComputerName $ServerActive_Unit -Port '10051' -WarningVariable bad_active| select ComputerName, TcpTestSucceeded)) 
                                {if 
                    ($bad_active) 
                        {Write-Output ("WARNING"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($bad_active -join ",")+";"+"id_77") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append}
                }
                }
            #check exe version, if not equal to repository version - update
            if (diff(get-childitem ($Global_Repo+"\ZABBIX_AGENTD.EXE") | Select-Object -ExpandProperty VersionInfo | Select-Object -ExpandProperty ProductVersion)(get-childitem ($Local_Repo+"\ZABBIX_AGENTD.EXE") | Select-Object -ExpandProperty VersionInfo | Select-Object -ExpandProperty ProductVersion))
                {
            Write-Output ("INFORMATION"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+"Updating ZABBIX_AGENTD.EXE - version is differ to"+($Global_Repo)+";"+"id_78") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
            Stop-Service -name (Get-WmiObject win32_service | Where-Object{($_.PathName -like '*zabbix*')}| Select-Object -ExpandProperty "Name") -ErrorAction SilentlyContinue

                try 

                            {
                            Get-Process | Where-Object{($_.Name -like '*zabbix*')} | Stop-Process -Force
                            Remove-Item -Path ($Local_Repo+"\ZABBIX_AGENTD.EXE") -Force | out-null
                            }

                        catch 

                            {
                            Write-Output ("Error"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($Error[0].Exception.Message )+";"+"id_79") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
                            Break
                            }

                    #copy from global repo
                    #here must be try\catch and break if fails
                    If (Test-Path $Global_Repo)
                        {

                        try

                            {
                            Copy-Item -Path ($Global_Repo+"\ZABBIX_AGENTD.EXE") -Destination $Local_Repo -Recurse
                            }

                        catch

                            {
                            Write-Output ("Error"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($Error[0].Exception.Message )+";"+"id_80") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
                            Break
                            }
                        }

                    else

                        {
                        Write-Output ("Error"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+"Cannot access to "+($Global_Repo)+";"+"id_81") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
                        Break
                        }
                    #output actions to the log
                    Write-Output ("INFORMATION"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($Local_Repo)+" ZABBIX_AGENTD.EXE has been updated from "+($Global_Repo)+";"+"id_82") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
                    #try register and start service
                    try

                        {
                        Start-Service -name (Get-WmiObject win32_service | Where-Object{($_.PathName -like '*zabbix*')}| Select-Object -ExpandProperty "Name")
                        #zabbix agent has been registered and running
                        Write-Output ("OK"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($ZBX_service.Name)+" service has been updated"+";"+"id_83") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append 
                        }

                    catch

                        {
                        #Cannot register and run zabbix_agentd.exe as service 
                        Write-Output ("Error"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($Error[0].Exception.Message )+";"+"id_84") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
                        Break
                        }
            
    }
            #check FQDN in ZABBIX_AGENTD.WIN.CONF, must be equal to real FQDN
            if ($Hostname = Get-Content ($Local_Repo+"\ZABBIX_AGENTD.WIN.CONF") | select-string -pattern "^Hostname=" -casesensitive |Select-Object -ExpandProperty line)
                {
            if (!(diff($Hostname)($Local_name)))
                {
                   try

                    {(Get-Content ($Local_Repo+"\ZABBIX_AGENTD.WIN.CONF")) -replace $Hostname, ("Hostname="+$Local_name) | Set-Content ($Local_Repo+"\ZABBIX_AGENTD.WIN.CONF")
                    Restart-Service -name (Get-WmiObject win32_service | Where-Object{($_.PathName -like '*zabbix*')}| Select-Object -ExpandProperty "Name")
                    Write-Output ("OK"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($ZBX_service.Name)+" FQDN has been set to ZABBIX_AGENTD.WIN.CONF"+";"+"id_85") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
                    }
    
                   catch

                    {
                    Write-Output ("WARNING"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($Error[0].Exception.Message )+";"+"id_86") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
                    }
                }

            }
            #If no FQDN defined - need to add
            else 
                {
                try

                    {
                    Add-Content -Value ("Hostname="+$Local_name) -Path ($Local_Repo+"\ZABBIX_AGENTD.WIN.CONF")
                    Restart-Service -name (Get-WmiObject win32_service | Where-Object{($_.PathName -like '*zabbix*')}| Select-Object -ExpandProperty "Name")
                    Write-Output ("OK"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($ZBX_service.Name)+" FQDN has been added to ZABBIX_AGENTD.WIN.CONF"+";"+"id_87") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
                    }
    
                catch

                    {Write-Output ("WARNING"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($Error[0].Exception.Message )+";"+"id_88") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append}
            }
            #check ZABBIX_AGENTD.WIN.CONF
            if (diff(get-content ($Global_Repo+"\ZABBIX_AGENTD.WIN.CONF"))(get-Content ($Local_Repo+"\ZABBIX_AGENTD.WIN.CONF") | select-string -notmatch -pattern "^Hostname=" -casesensitive))
                {
    Write-Output ("INFORMATION"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+"Updating ZABBIX_AGENTD.WIN.CONF - version is differ to"+($Global_Repo)+";"+"id_X10") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
        try

            {
            Set-Content -Value (get-content ($Global_Repo+"\ZABBIX_AGENTD.WIN.CONF")) -Path ($Local_Repo+"\ZABBIX_AGENTD.WIN.CONF")
            Add-Content -Value ("Hostname="+$Local_name) -Path ($Local_Repo+"\ZABBIX_AGENTD.WIN.CONF")
            Restart-Service -name (Get-WmiObject win32_service | Where-Object{($_.PathName -like '*zabbix*')}| Select-Object -ExpandProperty "Name")
            Write-Output ("OK"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+"ZABBIX_AGENTD.WIN.CONF is updated"+";"+"id_X11") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
            }
    
        catch

            {
            Write-Output ("Error"+";"+(Get-Date -Format "MM.dd.yyyy HH:mm")+";"+($env:computername)+";"+($Error[0].Exception.Message )+";"+"id_X12") | Out-File $Output\$Domain_name"_"zbx_agent.log -Append
            Break
            }

    }

        }
    }
#main end

Remove-Variable -Name * -Force -ErrorAction SilentlyContinue

