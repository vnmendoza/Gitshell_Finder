#~~~~~~~~~~~Noe's Awesome Script~~~~~~~~~~~
$PKG = ('2171AAC13F08109',
'2171AAC44B19300',
'2171AAC43L17931',
'2171AAC33I13166',
'2171AAC43L18067',
'2171AAC43L17633',
'2171AAC43L17822',
'2171AAC43L17635',
'2171AAC43L18216',
'2171AAC43L17737',
'2171AAC43L17747',
'2171AAC13F06028',
'2171AAC43L18073',
'2171AAC44B18909',
'2171AAC43L18171',
'2171AAC43L17702',
'2171AAC43L18126',
'2171AAC13F05537',
'2171AAC43L17642',
'2171AAC44B19110',
'2171AAC43L17970',
'2171AAC13E03378',
'2171AAC13F07323',
'2171AAC44B18703',
'2171AAC43L18069',
'2171AAC13G08203',
'2171AAC43L18215',
'2171AAC43L18027',
'2171AAC43L18191',
'2171AAC43L18212',
'2171AAC43L18022',
'2171AAC43L18033',
'2171AAC43L17906',
'2171AAC43L18208',
'2171AAC43L18210',
'2171AAC43L17459',
'2171AAC43L17869',
'2171AAC43L17842',
'2171AAC43L18123',
'2171AAC13F07751',
'2171AAC43L17655',
'2171AAC43L18124',
'2171AAC43L17744',
'2171AAC43L18190',
'2171AAC43L18019',
'2171AAC43L18172',
'2171AAC43L18015',
'2171AAC43L17886',
'2171AAC13G08643',
'2171AAC13F04003',
'2171AAC43L18193',
'2171AAC43L18220',
'2171AAC43L17991',
'2502ACA16F20399',
'2502ACA16F20412',
'2502ACA16F20379',
'2502ACA15K07273',
'2502ACA15K07012',
'2502ACA15K08256',
'2502ACA15K08041',
'2171AAC13F05578',
'2502ACA16E19677',
'2171AAC43L17674',
'2171AAC13F05571',
'2171AAC43L18169',
'2171AAC44B19348',
'1'
)
$msArrays = @(( "NULL","NULL"),
( "2171AAC13F08109","SVG 01"),
( "2171AAC44B19300","SVG 02"),
( "2171AAC43L17931","SVG 03"),
( "2171AAC33I13166","SVG 04"),
( "2171AAC43L18067","SVG 05"),
( "2171AAC43L17633","SVG 06"),
( "2171AAC43L17822","SVG 07"),
( "2171AAC43L17635","SVG 08"),
( "2171AAC43L18216","SVG 09"),
( "2171AAC43L17737","SVG 10"),
( "2171AAC43L17747","SVG 11"),
( "2171AAC13F06028","SVG 12"),
( "2171AAC43L18073","SVG 13"),
( "2171AAC44B18909","SVG 14"),
( "2171AAC43L18171","SVG 15"),
( "2171AAC43L17702","SVG 16"),
( "India","SVG 17"),
( "2171AAC43L18126","SVG 18"),
( "2171AAC13F05537","SVG 19"),
( "2171AAC43L17642","SVG 20"),
( "2171AAC44B19110","SVG 21"),
( "2171AAC13F05578","SVG 22"),
( "2171AAC13E03378","SVG 23"),
( "2171AAC13F07323","SVG 24"),
( "2171AAC43L18169","SVG 25"),
( "2171AAC44B18703","SVG 26"),
( "2171AAC13F05571","SVG 27"),
( "2171AAC43L18069","SVG 28"),
( "2171AAC13G08203","SVG 29"),
( "2171AAC43L17674","SVG 30"),
( "2171AAC43L17970","SVG 31"),
( "2171AAC43L18215","SVG 32"),
( "2171AAC43L18027","SVG 33"),
( "2171AAC43L18191","SVG 34"),
( "2171AAC43L18212","SVG 35"),
( "2171AAC43L18022","SVG 36"),
( "2171AAC43L18033","SVG 37"),
( "2171AAC43L17906","SVG 38"),
( "2171AAC43L18208","SVG 39"),
( "NULL","NULL"),
( "2171AAC43L18210","SVG 41"),
( "2171AAC43L17459","SVG 42"),
( "2171AAC43L17869","SVG 43"),
( "2171AAC43L17842","SVG 44"),
( "2171AAC13F07751","SVG 45"),
( "NULL","NULL"),
( "2171AAC43L18123","SVG 47"),
( "2171AAC43L17655","SVG 48"),
( "2171AAC43L18124","SVG 49"),
( "2171AAC43L17744","SVG 50"),
( "2171AAC43L18190","SVG 51"),
( "2171AAC43L18019","SVG 52"),
( "2171AAC44B19348","SVG 53"),
( "2171AAC43L18172","SVG 54"),
( "NULL","NULL"),
( "2171AAC43L18015","SVG 56"),
( "2171AAC43L17886","SVG 57"),
( "2171AAC13G08643","SVG 58"),
( "2171AAC13F04003","SVG 59"),
( "2171AAC43L18193","SVG 60"),
( "2171AAC43L18220","SVG 61"),
( "2171AAC43L17991","SVG 62"),
( "2502ACA16F20399","SVG 63"),
( "2502ACA16F20412","SVG 64"),
( "2502ACA16E19677","SVG 65"),
( "2502ACA16F20379","SVG 66"),
( "2502ACA15K07273","SVG 67"),
( "2502ACA15K07012","SVG 68"),
( "2502ACA15K08256","SVG 69"),
( "2502ACA15K08041","SVG 70"),
( "1","SVG 100")



)



$Date='01/03/2018'              #Enter Date to Seach for (Example: '04/05/2016' )
$FirstRun='yes'                   #Enter Yes if this is the first run of the day or if you closed Powershell ISE, it will load up the master XML.
                                 #This process takes a while so try not to close it after doing it once 
$path="C:\Users\amejia\Desktop\Incentive Range"                  
#
#
#   ***PRESS F5 or the PLAY button up top to RUN***
#
if ($FIRSTRUN -like "Yes") {
    write-host "Parsing through Master XML folder"
    $xmls_full += Get-ChildItem ("P:\CYIENT 2015\PolesBackup\Pole Back Up Data\Master XML\*.xml") #Parsing through the Master takes forever, preferably you only want to do it once and use that data for time searches
    }



$d1 = [DateTime]::Parse($Date)

$d2=$d1.AddDays(1)
$d1=$d1.ToShortDateString()
$d2=$d2.ToShortDateString()
#var
    $htaxmls= @()           
$completeXMLS= @()
$completeXMLSRefield= @()
$completeIDs=@()
$htaIDs =@()
$completeIDsRefield=@()
$patsCodes= @()
$pkgNums = @()
$htaPKGnums =@()
$exportMe = @()
$you = @()
$you2 = @()
$intermediateArr = @()
$users = @()
$onlyCount = @()


$start = Get-Date
$todate = $date.Replace("/",".")
$j=0
$table=@"
User, MID, SVG, Count, HTA
"@
$table | Set-Content ($path + "\APoles_"+$todate+".csv")
 



write-host "Running Date Search for"$d1
$xmls += $xmls_full | Where-Object {($_.LastWriteTime -gt $d1 -and $_.LastWriteTime -lt $d2 )} #time stamp search
$total = $xmls.Count

write-host "Parsing through XMLs (finding complete poles)"

foreach($xml in $xmls) 
{
    # Progress Tracking
	$j++
	$prct = [Math]::Round((($j / $total) * 100.0), 2)
 
	$elapsed = (Get-Date) - $start
	$totalTime = ($elapsed.TotalSeconds) / ($prct / 100.0)
	$remain = $totalTime - $elapsed.TotalSeconds
	$eta = (Get-Date).AddSeconds($remain)
	
	# Display
	
	Write-Progress -Activity "ETA $eta" -Status "$prct" -PercentComplete $prct
	
	# Operation
    $xmlContent = [xml](get-content $xml)

    #New Solution
    if ($xmlContent.NewDataSet.POLE)
    {
        $patsCode = $xmlContent.NewDataSet.POLE.STATUS_CODE          #PATS data
        $pkgNum   = $xmlContent.NewDataSet.POLE.ASSIGNED_TO          #PKG data
        $POLEID   = $xmlContent.NewDataSet.POLE.POLE_ID              #Pole ID
        $user     = $xmlContent.NewDataSet.POLE.FUSERID              #User ID
        $htaReason = $xmlContent.NewDataSet."CIP_SH.HTA".HTA_REASON  #HTAReason
        if ($xmlContent.NewDataSet.POLE.FCCOMPDATE)
        {
            $DateMod   =$xmlContent.NewDataSet.POLE.FCCOMPDATE
            $Seperator = "T"
            $shortDate = $DateMod.split($Seperator)
        }
     
        if((($patsCode -eq "FIELD COLLECTION COMPLETED") ) -and ($PKG -contains $pkgNum) ) #finds all completed and specific pkg num. Comment out pkg num if you want all the poles completed
        {          
            $completeXMLS += $xml #adds the current xml to an xml that only contains completed poles
            $completeIDs += ($POLEID + "," + $pkgnum + "," + $shortDate[0] )
            $cUser = $user            
        }
        if((($patsCode -eq "HTA") ) -and ($PKG -contains $pkgNum) -and ($htaReason -ne "CAN'T LOCATE")) #finds all completed and specific pkg num. comment out pkg num if you want all the poles completed
        {          
            $htaXMLS += $xml #adds the current xml to an xml that only contains completed poles
            $htaIDs += ($POLEID + "," + $pkgnum + "," + $shortDate[0] )
            $hUser= $user         
        }     
    }
    
    #BAU/ Old Solution
    else 
    {
        $patsCode = $xmlContent.NewDataSet.CALCPOINT.PATS #PATS data
        $pkgNum = $xmlContent.NewDataSet.CALCPOINT.ASGN_TO #PKG data
        $POLEID =$xmlContent.NewDataSet.CALCPOINT.POLE_ID
        if($xmlContent.NewDataSet.CALCPOINT.DATEMODIFIED)
        {
            $DateMod =$xmlContent.NewDataSet.POLE.FCCOMPDATE
            $Seperator = "T"
            $shortDate = $DateMod.split($Seperator)
        }
        if((($patsCode -eq 6) ) -and ($PKG -contains $pkgNum) ) #finds all completed and specific pkg num. comment out pkg num if you want all the poles Pass 2 complete
        {
            $completeXMLS += $xml #adds the current xml to an xml that only contains completed poles
            $completeIDs += ($POLEID + "," + $pkgnum + "," + $shortDate[0])  
            $cUser = $user         
        } 
    
        elseif((($patsCode -eq 15)) -and ($pkgnum -eq ($PKG)) ) #finds all completed and specific pkg num. comment out pkg num if you want all the poles Rework Complete
        {
            $completeXMLSRefield += $xml #adds the current xml to an xml that only contains refield poles
            $completeIDsRefield += ($POLEID + "," + $DateMod)
        }
    }
    $k=0
    foreach($msArr in $msArrays)
    {        
        $svg = ("SVG " + $k.ToString("00"))
        $lookUP = ($user, $pkgNum, $svg)
        if($user -notlike "*-QC")
        {
            if($msArr[0] -eq $pkgNum)
            {
                $msArrays[$k] = ,$user + $msarr
                $users += $user
            
                #        
                #$ass = @($user,$pkgNum,$svg)
                #$you += ,$ass

            }

            elseif(($users -notcontains $user) -and ($msArr.Contains($pkgNum)))
            {
                $msArrays += ,($user, $pkgNum, $svg)
                $users += $user
            }
        }
        $k++
    }    
}


 

write-host "counting poles"
$completeIDs = $completeIDs | select -uniq
$htaids = $htaids | select -uniq
$completeIDsRefield = $completeIDsRefield | select -uniq
for($i = 0; $i -le ($completeIDs.Length-1); $i++)
    {
    $pkgnums += $completeIDs[$i].Split(',')[1]      
    
    }
for($i = 0; $i -le ($htaIDs.Length-1); $i++)
    {
    $htapkgnums += $htaIDs[$i].Split(',')[1]        
    
    }


$pkgCount = $pkgnums | group | % { $h = @{} } { $h[$_.Name] = $_.Count } { $h }
$htapkgCount = $htapkgnums | group | % { $h = @{} } { $h[$_.Name] = $_.Count } { $h }

$k=0
$j = 0
foreach($u in $msArrays)
{    
    foreach ($pack in $pkgCount.GetEnumerator())
    {
        if($pack.Key -eq $u[1])
        {
            $onlyCount += ,($msArrays[$k] + $pack.Value)
            $j++            
        }        
    }
    foreach ($htapack in $htapkgCount.GetEnumerator())
    {
        if($htapack.Key -eq $u[1])
        {
            $onlyCount[($j - 1)] += ($htapack.Value)            
        }        
    } 
    $k++  
}
#foreach 
#$intermediateArr[0] | select -Unique
foreach ($export in $onlyCount)
{
    $eUser = $export[0]
    $eMID = $export[1]
    $eSVG = $export[2]
    $eCount = $export[3]
    $eHTA = $export[4]
    Add-Content ($path + "\APoles_"+$todate+".csv") "$eUser,$eMID,$eSVG,$eCount,$eHTA"  
}
write-host "Creating Files"
#xml filenames
$completeXMLS| Select-object name | Out-File ($path + "\FullFilename_" + $todate +".txt")
$htaXMLS| Select-object name | Out-File ($path + "\FullFilename_" + $todate + ".txt")
$completeXMLSRefield| Select-object name | Out-File ($path + "\FullFilename_Refield_" + $todate +".txt")
#Pole IDs
$completeIDs | Out-File ($path + "\Pole_IDs_" + $todate +".txt")
$htaIDs | Out-File ($path + "\Pole_htaIDs_" + $todate +".txt")
import-csv ($path + "\Pole_IDs_" + $todate +".txt") -delimiter "," | export-csv ($path + "\Pole_IDs_" + $todate +".csv")
$completeIDsRefield | Out-File ($path + "\Pole_IDs_Refield_" + $todate +".txt")

$pkgCount | Out-File ($path + "\SyncReportAmounts_" + $todate +".txt")
$pkgnums | Out-File ($path + "\allpkkg_" + $todate +".txt")


$msArrays | Out-File ($path + "\ult_IDs_" + $todate +".txt")
import-csv ($path + "\ult_IDs_" + $todate +".txt") -delimiter "," | export-csv ($path + "\ult_IDs_" + $todate +".csv")



#export to a csv file




write-host "done"
[Console]::Beep(130, 100)
[Console]::Beep(262, 100)
[Console]::Beep(330, 100)
[Console]::Beep(392, 100)
[Console]::Beep(523, 100)
[Console]::Beep(660, 100)
[Console]::Beep(784, 300)
[Console]::Beep(660, 300)
[Console]::Beep(146, 100)
[Console]::Beep(262, 100)
[Console]::Beep(311, 100)
[Console]::Beep(415, 100)
[Console]::Beep(523, 100)
[Console]::Beep(622, 100)
[Console]::Beep(831, 300)
[Console]::Beep(622, 300)
[Console]::Beep(155, 100)
[Console]::Beep(294, 100)
[Console]::Beep(349, 100)
[Console]::Beep(466, 100)
[Console]::Beep(588, 100)
[Console]::Beep(699, 100)
[Console]::Beep(933, 300)
[Console]::Beep(933, 100)
[Console]::Beep(933, 100)
[Console]::Beep(933, 100)
[Console]::Beep(1047, 400)
    

#Read-Host -Prompt "Done. Press Enter to exit. This won't close Powershell ISE"

    