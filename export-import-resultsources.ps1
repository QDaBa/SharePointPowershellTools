#-----------------------------------------------------------------------------  
# Name:             export-import-resultsources.ps1   
# Description:      This script has two switch to:
#                    - export a list of result sources (optionally filtered) 
#                    - import a list of result sources
# Usage:            Run the script passing paramters ServiceApp, EnterpriseSearchOwnerLevel Import|Export and Filter  
# By:               Quoc Dang Banh, inspired by Tobias Lekman (tobias@lekman.com) https://blog.lekman.com/2015/08/script-to-import-export-compare-and.html
#                   and Riccardo Celesti blog.riccardocelesti.it https://gallery.technet.microsoft.com/Powershell-script-to-09ffa974
#----------------------------------------------------------------------------- 


[CmdletBinding(DefaultParameterSetName = "Export")] 

Param([Parameter(Mandatory=$true)] 
      [String]$serviceapp,
      [Parameter(Mandatory=$true)] 
      [String]$enterpriseSearchOwnerLevel,      
      [Parameter(Mandatory=$false,ParameterSetName='Export')]      
      [String]$filter,
      [Parameter(ParameterSetName='Import')]
      [switch]$import,
      [Parameter(ParameterSetName='Export')]
      [switch]$export   
)
if ((gsnp MIcrosoft.SharePoint.Powershell -ea SilentlyContinue) -eq $null){
    asnp Microsoft.SharePoint.Powershell -ea Stop
}
$logfile = ".\export-resultsources.csv"
$importlog = ".\import-resultsources.log"
$importlogerr = ".\import-resultsources-errors.log"

if ((Get-SPEnterpriseSearchServiceApplication $serviceapp -ea SilentlyContinue) -eq $null){
    Write-Host "Enterprise Search Service Application $serviceapp has not been found" -ForegroundColor Red
    exit
} else {
    $ssa = Get-SPEnterpriseSearchServiceApplication $serviceapp
    Write-Host "Enterprise Search Service Application $serviceapp has been found" -ForegroundColor Green
}

$owner = Get-SPEnterpriseSearchOwner -Level $enterpriseSearchOwnerLevel

if ($export){

    if ((Get-ChildItem -Name $logfile -ea SilentlyContinue) -ne $null){
        Clear-Content $logfile
        ac $logfile "Id,Name,Owner,ProviderId,QueryTemplate";
    } else {
        ac $logfile "Id,Name,Owner,ProviderId,QueryTemplate";
    }
    if ($filter -ne $null){
        Get-SPEnterpriseSearchResultSource -SearchApplication $ssa -Owner $owner | ?{$_.Name -like "$($filter)*"} | %{        
            $rs = $_;            
            ac $logfile "$($rs.Id),$($rs.Name),$($rs.Owner),$($rs.ProviderId),$($rs.QueryTransform.QueryTemplate)"            
        }

    } else {
        Get-SPEnterpriseSearchResultSource -SearchApplication $ssa -Owner $owner | %{        
            $rs = $_;                       
            ac $logfile "$($rs.Id),$($rs.Name),$($rs.Owner),$($rs.ProviderId),$($rs.QueryTransform.QueryTemplate)"
        }
    }    
}

if ($import){
    if ((Get-ChildItem $logfile -ea SilentlyContinue) -eq $null){
        Write-Host "The export file has not been found" -ForegroundColor Red
        exit
    } else {
        Import-Csv $logfile -Delimiter "," | %{
            $rs = $_;
            if ((Get-SPEnterpriseSearchResultSource -SearchApplication $ssa -Owner $owner | ?{$_.Name -eq $rs.Name} -ea SilentlyContinue) -eq $null){
                try {
                    $newrs = New-SPEnterpriseSearchResultSource -SearchApplication $ssa -Owner $owner -Name $rs.Name -ProviderId $rs.ProviderId -QueryTemplate $rs.QueryTemplate  -EA Stop
                    ac $importlog "$(Get-Date),$($rs.Name),ResultSource,"
                    
                } catch {
                    Write-Host "Something went wrong :) $($Error[0].Exception.Message)" -ForegroundColor Red
                    ac $importlog "$(Get-Date),$($rs.Name),[ERR],"
                    ac $importlogerr "$(Get-Date),$($rs.Name),[ERR],ResultSource,$($Error[0].Exception.Message)"
                }
            } else {
                Write-Host "ResultSource '$($rs.Name)' already exists" -ForegroundColor Red
                ac $importlog "$(Get-Date),$($rs.Name),[WRN],ResultSource,ResultSource already exists"
            }
        }
    }
}