# Zabbix_agent update and install
 
 
Does: 


1. Check, if zabbix agent service is installed and running. If not - fixes it.

2. Check, if firewall rules are set. If not - script will create rules.

3. Check, if active servers is reacheble for active zabbix agent. If not - out warning to the log.

4. Check, if zabbix agent exe version is equal to repo version. If not - fixes it.

5. Check, if zabbix agent conf has all string matched to repo conf version(Hostname string is excluded from matching). If not - fixes it.

6. Check, if low register FQDN is set correctly in zabbix agent conf. If not - fixes it.


Versions:


1. Zabbix_agent_remote - runs remotely via invoke-command for each windows server in AD. AD module need to be installed to collect servers list. Servers accounts in AD need to be enabled for delegation for any service via Kerberos auth. 

2. Zabbix_agent_local - runs localy on server. Can be distributed via GPO.


Basic scenarios to handle:


1. no zabbix installed

2. no firewall rule is set

3. no binares, service exist

4. no service, binares exist

5. service exist, binares are broken

6. binares exist, service is broken path

7. binares and service are broken

8. no binares, service are broken

9. local repo is empty, service exist

10. no service, binares are broken

11. service ok, config is broken

12. zabbix agent version is not equal to repo version

13. zabbix conf is not equal to repo version, Hostname string is excluded from compare


Conditions:


-Windows 2012 R2 (for Windows 2008 R2 cannot create firewall rule and check connectivity between host and zabbix server)

-Default path for zabbix agent need to be c:\Program Files\zabbix40

-Repository need to placed in NETLOGON\zabbix40
