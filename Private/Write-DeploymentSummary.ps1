function Write-DeploymentSummary {
    param (
        [string] $FileName,
        [string] $TableName,
        [string] $SiteCode,
        [int] $NumberOfDays,
        $LogFile,
        [string] $ServerName,
        $ContinueOnError = $true
    )
    Write-Log -Message "[function: write-deploymentsummary]" -LogFile $logfile
    $query = "SELECT SoftwareName,AssignmentID,CollectionName,CollectionID,DeploymentTime,
        CreationTime,ModificationTime,
        case 
            when (featuretype = 1) then 'Application'
            when (featuretype = 2) then 'Program'
            when (featuretype = 3) then 'MobileProgram'
            when (featuretype = 4) then 'Script'
            when (featuretype = 5) then 'SoftwareUpdate'
            when (featuretype = 6) then 'Baseline'
            when (featuretype = 7) then 'TaskSequence'
            when (featuretype = 8) then 'ContentDistribution'
            when (featuretype = 9) then 'DistributionPointGroup'
            when (featuretype = 10) then 'DistributionPointHealth'
            when (featuretype = 11) then 'ConfigurationPolicy'
            when (featuretype = 28) then 'AbstractConfigurationItem'
        end as FeatureType,
        SummaryType,
        case 
            when (DeploymentIntent = 1) then 'Install'
            when (DeploymentIntent = 2) then 'Uninstall'
            when (DeploymentIntent = 3) then 'Preflight'
            end as DeployIntent,
        EnforcementDeadline,
        NumberTotal as Total,
        NumberSuccess as Success,
        NumberErrors as Failed,
        NumberInProgress as InProgress,
        NumberUnknown as Unknown,
        NumberOther as Other,
        SummarizationTime,
        ProgramName,
        PackageID
    FROM vDeploymentSummary
    WHERE FeatureType <> 5
    ORDER BY SoftwareName"
    $ds = @(Invoke-DbaQuery -SqlInstance $ServerName -Database $SQLDBName -Query $query -ErrorAction SilentlyContinue)
    if ($null -eq $blist) { return }
    $Fields = @("SoftwareName","AssignmentID","CollectionName","CollectionID","DeploymentTime","CreationTime","ModificationTime","FeatureType","SummaryType","DeployIntent","EnforcementDeadline","Total","Success","Failed","InProgress","Unknown","Other","SummarizationTime","ProgramName","PackageID")
    $dsDetails = New-CmDataTable -TableName $tableName -Fields $Fields
    foreach ($b in $blist) {
        $row = $dsDetails.NewRow()
        $row.SoftwareName = $ds.SoftwareName
        $row.AssignmentID = $ds.AssignmentID
        $row.CollectionName = $ds.CollectionName
        $row.CollectionID = $ds.CollectionID
        $row.DeploymentTime = $ds.DeploymentTime
        $row.CreationTime = $ds.CreationTime
        $row.ModificationTime = $ds.ModificationTime
        $row.FeatureType = $ds.FeatureType
        $row.SummaryType = $ds.SummaryType
        $row.DeployIntent = $ds.DeployIntent
        $row.EnforcementDeadline = $ds.EnforcementDeadline
        $row.Total = $ds.Total
        $row.Success = $ds.Success
        $row.Failed = $ds.Failed
        $row.InProgress = $ds.InProgress
        $row.Unknown = $ds.Unknown
        $row.Other = $ds.Other
        $row.SummarizationTime = $ds.SummarizationTime
        $row.ProgramName = $ds.ProgramName
        $row.PackageID = $ds.PackageID
        $dsDetails.Rows.Add($row)
    }
    , $dsDetails | Export-CliXml -Path ($filename)
}