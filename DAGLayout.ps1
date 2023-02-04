<#--------------
# Script Info
#--------------

# Script Name                :         Exchange DAG Database Distribution Table
# Script Version             :         1.1
# Author                     :         Ammar Hasayen
# Blog                       :         http://ammarhasayen.wordpress.com
# Description                :         This script will Query all DAGs in your environment and will create table with the following information
                                       - list DB copies hosted in each server
                                       - list the activation preference for each database in nice table presentation
                                       - Highlight the preferred server for the database to be mounted on according to the Activation Preference (Yellow Highlight)
                                       - Highlight the current server that the database is actually mounted on ( Red highlight if NOT the preferred)
                                       - Summary for :
                                                1. Total Number of databases copies on each server
                                                2. Ideal number of databases that should be mounted on each server
                                                3. Actual number of databases mounted on each server

                                        The script will send email with those info at the end

# Slight Modification by Sam and Bernie Chouinard: 
- added color when database is mounted correctly on Activation Preference = 1 (dark green)
- added send e-mail info as optional parameters with default values
- generating different HTML file each run with date and time stamp (might need to cleanup directory with lots of HTM if need be)


#--------------
# Script Requirement
#--------------

# -Open Exchange PowerShell with account that has Exchange View Only Administrator
 



#--------------
# Notes
#--------------

The script will automatically create an HTML file in the (Get-Location) path from where you run the script.

Total Copies: represent the total number of database copies (Active and Passive) that are mounted on the server
Ideal Mounted DB Copies : According to the Activation Preference, how many databases should be mounted on the server.In other words, how many databases have this server with Activation preference = 1
Actual Mounted DB Copies : How many databases actually mounted on this server
Yellow cells, represent the server on which the database is mounted and it happens that it is mounted on the server with Activation preference = 1
Red cells, represent the server on which the database is mounted and it happens that it is mounted on the server with Activation preference not equal 1
>Green cells are database copy locations with activation preference

#--------------
# Script Start
#--------------
#>

[CmdletBinding()]
Param (
   [string]$eMailSender="no-reply@canadadrey.ca",
   [string]$eMailRecipient="samdrey@canadadrey.ca",
   [string]$eMailServer="mail.canadadrey.ca"
)

$CurrentScriptLocation = Get-Location
$File = "\DAG_DB_Layout$(Get-Date -Format 'yyyy-MM-dd_hhmmss').htm"


New-Item `
 -ItemType file `
    -Name $File `
      -path $CurrentScriptLocation  `
            -Force

$filename = "$($CurrentScriptLocation)"+ $File


#--------------
# Script Functions
#--------------
Function sendEmail 
{ param($from,$to,$subject,$smtphost,$htmlFileName) 


        $msg = new-object Net.Mail.MailMessage
        $smtp = new-object Net.Mail.SmtpClient($smtphost)
        $msg.From = $from
        $msg.To.Add($to)
        $msg.Subject = $subject
        $msg.Body = Get-Content $htmlFileName 
        $msg.isBodyhtml = $true 
        $smtp.Send($msg)
       

} 


function _GetDAG
{
	param($DAG)
	@{Name			= $DAG.Name.ToUpper()
	  MemberCount	= $DAG.Servers.Count
	  Members		= [array]($DAG.Servers | % { $_.Name })
	  Databases		= @()
	  }
}



function   _GetDB 
{
  param($Database)

  $DB_Act_pref = $Database.ActivationPreference
  $Mounted = $Database.Mounted
  $DB_Act_pref = $Database.ActivationPreference
  [array]$DBHolders =$null 
		( $Database.Servers) |%{$DBHolders  += $_.name}





     @{Name						= $Database.Name
	  ActiveOwner				= $Database.Server.Name.ToUpper()	 
	  Mounted                   = $Mounted
	  DBHolders			        = $DBHolders
	  DB_Act_pref               = $DB_Act_pref 	  
	  IsRecovery                = $Database.Recovery
	  }

}


function _GetDAG_DB_Layout
{
	param($Databases,$DAG)

	    $WarningColor                      = "#FF9900"
		$ErrorColor                        ="#980000"
		$BGColHeader                       ="#000099"
		$BGColSubHeader                    ="#0000FF"
		[Array]$Servers_In_DAG             = $DAG.Members
		
    $Output2 ="<table border=""0"" cellpadding=""3"" width=""50%"" style=""font-size:8pt;font-family:Arial,sans-serif"">
	<col width=""5%"">
	<colgroup width=""25%"">"
	$Servers_In_DAG | Sort-Object | %{$Output2+="<col width=""3%"">"}
	$Output2 +="</colgroup>"
	$ServerCount = $Servers_In_DAG.Count
	
	$Output2 += "<tr bgcolor=""$($BGColHeader)""><th><font color=""#FFFFFF"">DatabaseCopies</font></th>	
	<th colspan=""$($ServerCount)""><font color=""#FFFFFF"">Mailbox Servers in $($DAG.name)</font></th>	
	</tr>"
	$Output2+="<tr bgcolor=""$($BGColSubHeader)""><th></th>"
	$Servers_In_DAG|Sort-Object | %{$Output2+="<th><font color=""#FFFFFF"">$($_)</font></th>"}
	
	$Output2 += "</tr>"
	
	#writing table content
	$AlternateRow=0
		foreach ($Database in $Databases)
	{
	$Output2+="<tr "
	if ($AlternateRow)
					{
						$Output2+=" style=""background-color:#dddddd"""
						$AlternateRow=0
					} else
					{
						$AlternateRow=1
					}
		
		$Output2+="><td><strong>$($database.name)</strong></td>"
		
		#copies
		
							$DatabaseServer   = $Database.ActiveOwner
							$DatabaseServers  = $Database.DBHolders
		$Servers_In_DAG|Sort-Object| 
			%{ 
									 $ActvPref =$Database.DB_Act_pref
									 $server_in_the_loop = $_
									 $Actv = $ActvPref  |where {$_.key -eq  $server_in_the_loop}
									 $Actv=  $Actv.value
									 $ActvKey= $ActvPref |Where {$_.value -eq 1}
									 $ActvKey = 	 $ActvKey.key.name
									  
									  
							$Output2+="<td"
							
								if (  ($DatabaseServers -contains $_) -and ( $_ -like $databaseserver)  )
										{
											if (  $ActvKey -like $databaseserver  )
											{$Output2+=" align=""center"" style=""background-color:#008000""><font color=""#000000f""><strong>$Actv</strong></font> "}
											else
											{$Output2+=" align=""center"" style=""background-color:#FB0B1B""><strong><font color=""#FFFFFF"">$Actv</strong></font> "}
										
										}
					
							
										elseif ($DatabaseServers -contains $_)
									{
								
									
									$Output2+=" align=""center"" style=""background-color:#00FF00"">$Actv "							 
									}
									else
								{ $Output2+=" align=""center"" style=""background-color:#dddddd"">"	}
								 
			
			}
				
		
		
		$Output2+="</tr >"
		}
		
	$Output2+="<tr></tr><tr></tr><tr></tr>"
	
	
	#Total Assigned copies
	
	$Output2 += "<tr bgcolor=""#440164""><th><font color=""#FFFFFF"">Total Copies</font></th>"
	
	$Servers_In_DAG|Sort-Object| 
			%{ 
		
								
								$this = $EnvironmentServers[$_]
									$Output2 += "<td align=""center"" style=""background-color:#E0ACF8""><font color=""#000000""><strong>$($this.DBCopyCount)</strong></font></td>"	
						
		}
	$Output2 +="</tr>"
	#Copies Assigned Ideal
	
	$Output2 += "<tr bgcolor=""#DB08CD""><th><font color=""#FFFFFF"">Ideal Mounted DB Copies</font></th>"
	
	$Servers_In_DAG|Sort-Object| 
			%{ 
			foreach ($this in $My_Hash_3.GetEnumerator())
						{				
									if ($this.key  -like $_)
									{$Output2 += "<td align=""center"" style=""background-color:#FBCCF9""><font color=""#000000""><strong>$($this.value)</strong></font></td>"}		
						}
		}
	$Output2 +="</tr>"
	
# Copies Actually Assigned
	
	$Output2 += "<tr bgcolor=""#440164""><th><font color=""#FFFFFF"">Actual Mounted DB Copies</font></th>"
	
	$Servers_In_DAG|Sort-Object| 
			%{ 
		
								
								$this = $EnvironmentServers[$_]
									$Output2 += "<td align=""center"" style=""background-color:#E0ACF8""><font color=""#000000""><strong>$($this.DBCopyCount_Assigned)</strong></font></td>"	
						
		}
	$Output2 +="</tr>"	
		
$Output2


}



#Gets all Mailbox Servers with extra info and return a hashtable
function _GetMailboxServerInfo
{
     param($Server, $databases)

     [int]$DBCopyCount           = -1 # This is considered not initialized
     [int]$DBs_Mountedcount       = 0 

     #Getting DBs mounted on this server
     $DBs_Mounted                      = @($databases | Where {$_.Server -ieq $Server})
     $DBs_MountedCount                 =  $DBs_Mounted.count

     #Getting DB copies on this server
     $MailboxServer                         = Get-MailboxServer $Server
     [array]$MailboxServer_DB_Copies        = @()     
     Try{$MailboxServer_DB_Copies           = $MailboxServer  | 
                                                Get-MailboxDatabaseCopyStatus -ErrorAction SilentlyContinue}
                                                  
                                                  Catch{}


    if ($MailboxServer_DB_copies) #if the server has copies
            {
              $DBCopyCount           = $MailboxServer_DB_copies.Count  
            }


     #Return hashtable
     @{ Name = $server
      DBCopyCount = $DBCopyCount
      DBCopyCount_Assigned = $DBs_MountedCount
      }


    

}


#Gets all Mailbox Servers in the organization
function _GetMailboxServers
{

$ExchServers = Get-ExchangeServer |Where {$_.ServerRole -Contains "Mailbox"}

Return $ExchServers 

}


#--------------
# Script Code
#--------------



#--------------------START Global Variables---------------------------

$Databases           = [array](Get-MailboxDatabase -Status) 
$My_Databases        = $Databases |where {$_.Recovery -eq $false}
$DAGs                = [array](Get-DatabaseAvailabilityGroup)
$SRV                 = @() #Holds name of Exchange Mailbox Servers participating in DAGs
$EnvironmentServers  = @{} # Hashtable to hold Exchange Mailbox Server custom info 
$EnvironmentDAGS     = @() # Hashtable to hold DAG custom info


#--------------------END Global Variables---------------------------



#--------------------START Collecting DAG Info-----------------------



if ($DAGs) {
        Foreach ($DAG in $DAGS)
                {
                $EnvironmentDAGS += _GetDAG $DAG
                }
            }

#-------------------- END Collecting DAG Info------------------------


#-------------------- START Collecting DB Info-----------------------

for ($i=0; $i -lt $My_Databases.Count; $i++){

     $database = _GetDB $My_Databases[$i]


    for ($j=0; $j -lt $EnvironmentDAGS.Count; $j++)
			{
				if ($EnvironmentDAGS[$j].Members -contains $Database.ActiveOwner)
				{
					$EnvironmentDAGS[$j].Databases += $Database
				}
			}
}

#-------------------- END Collecting DB Info-----------------------


#-------------------- START Collecting Exchange Server Info---------

#collect DAG Exchange Servers

Foreach($DAG in $DAGS){
 Foreach ($Server in $DAG.Servers){
  $SRV+=$Server.name}}
   
 
foreach ($server in $SRV)

        { $SRV_Info  =_GetMailboxServerInfo $server $My_Databases

         $EnvironmentServers.Add($SRV_Info.Name, $SRV_Info) 

         }


#-------------------- END Collecting Exchange Server Info-----------


#-------------------- Start Creating HTML Table-----------


#Hold Info in temp hashtables
$My_Hash_1 = $null
$My_Hash_1 = @{}

$My_Hash_2 = $null
$My_Hash_2 = @{}

$myobjects =$null
$myobjects = @()

$My_Hash_3 =$null
$My_Hash_3 = @{}





        foreach ($My_database in $My_databases)

                {
	                $Var1 =  $My_database.activationPreference |where{$_.value -eq 1}
	                $Var2 = $Var1.key.name
	                $My_Hash_1.add($My_Database.Name,$Var2)
                }

        foreach ($Var3 in $My_Hash_1.GetEnumerator())
                {
                $objx = New-Object System.Object
                $objx | Add-Member -type NoteProperty -name Name -value $Var3.key
                $objx | Add-Member -type NoteProperty -name count -value $Var3.value
                $myobjects +=     $objx
                }


$mydata = $myobjects |Group-Object -Property count

        foreach ($counting in $mydata)
                {

                $My_Hash_3.add($counting.name,$counting.count)
                }

        foreach ($Server in $SRV)
        {
                if(!( $My_Hash_3.ContainsKey($Server)))
                {
                  $My_Hash_3.Add($Server, 0)
                }

        }




$Output ="<html>
<body>
<font size=""1"" face=""Arial,sans-serif"">
<h3 align=""center"">DAG Database Copies layout</h3>
<h5 align=""center"">Generated $((Get-Date).ToString())</h5>
</font>"




        foreach ($DAG in $EnvironmentDAGS )
                {
	                if ($DAG.Membercount -gt 0)
	                        {
		                        # Database Availability Group Header
		                        $Output +="<table border=""0"" cellpadding=""3"" width=""50%"" style=""font-size:8pt;font-family:Arial,sans-serif"">
		                        <col width=""20%""><col width=""10%""><col width=""70%"">
		                        <tr align=""center"" bgcolor=""#FC8E10""><th><font color=""#FFFFFF"">Database Availability Group Name</font></th><th><font color=""#FFFFFF"">Member Count</font></th>
		                        <th><font color=""#FFFFFF"">Database Availability Group Members</font></th></tr>
		                        <tr><td>$($DAG.Name)</td><td align=""center"">
		                        $($DAG.MemberCount)</td><td>"
		                        $DAG.Members | % { $Output+="$($_) " }
		                        $Output +="</td></tr></table>"
		
		
		
		                        # Get Table HTML

                            $Output += _GetDAG_DB_Layout -Databases $DAG.Databases -DAG $DAG
	                       

                            }
	
                }

$Output += "</table>"

$Output+="</body></html>";
Add-Content $filename $Output


#-------------------- END Creating HTML Table-----------


#-------------------- START Send Email-------------------



send-mailmessage -from $eMailSender -to $eMailRecipient -subject "DAG Layout Report _$(Get-Date -f 'yyyy-MM-dd')" -SMTPServer $eMailServer -Attachments $fileName

#NOTE: if you receive a 5.7.1 Message rejected as spam by Content Filtering you can run the below:
# Set-ContentFilterConfig -BypassedSenderDomains yourdomain

#-------------------- END Send Email---------------------


#--------------
# END SCRIPT
#--------------
