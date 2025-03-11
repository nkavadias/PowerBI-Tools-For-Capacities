################################################################################################################################################################################
# Script executes the load test
# If run with no parameters, it will execute the load test found in each subfolder:
#  .\Run_Load_Test_Only.ps1
# If run with parameters, it will only execute the subfolders specified:
#  .\Run_Load_Test_Only.ps1 "DemoLoadTest1" "DemoLoadTest2"
# It pauses 5 seconds between opening each window so as not to overload the client machine CPU during initiation of the load test.
# Once load test windows are opened, it waits for the user to press enter again at which point it closes *all* open Chrome windows for the current user. Note it does not just close Chrome windows it opens.
# Run this script as administrator
################################################################################################################################################################################
[CmdletBinding()]
Param(
  [Parameter(Mandatory = $false )]     [switch] $UseDocker,
  [Parameter(Mandatory = $false )]     [string] $Instances,
  [Parameter(Mandatory = $false )]     [string] $TestNameDirectoryFilter
  
)

Function dockerRun {

    # Check if Docker is installed
if (-not (Get-Command docker -ErrorAction SilentlyContinue)) {
    Write-Host "Docker is not installed or not in PATH. Please install Docker Desktop first." -ForegroundColor Red
    exit 1
}
# Pull the Docker image
Write-Host "Pulling Docker image $image..."
docker pull $image


}

function CheckForSeleniumBIModule
{
    $module = Get-Module -ListAvailable -Name Selenium 
    if($module -eq $null)
    {
        Write-Host "Selenium module is not installed. Please install the module before proceeding further." -ForegroundColor Yellow
        try {
        Write-Host "Trying to install Selenium module" -ForegroundColor Yellow    
        Install-Module -Name Selenium  
        }
        catch {
            Write-Host "Failed to install Selenium. Please install the module manually and then proceed further." -ForegroundColor Red
            exit
        }
    } 
}

$image = "selenium/standalone-chrome"
$containerName = "PBILoadTest"

$htmlFileName = 'RealisticLoadTest.html'
$workingDir = $pwd.Path
"This script finds all subdirectories with $htmlFileName files and runs a specifies number of instances of each."
$instances = [int] $(Read-Host -Prompt 'Enter number of instances to initiate for each report')
$numberOfPhysicalCores = (Get-ciminstance –class Win32_processor).NumberOfCores;
if ($numberOfPhysicalCores.Length)
{
    #if computer has multiple sockets, then sum the array
    $numberOfPhysicalCores = ($numberOfPhysicalCores | Measure-Object -Sum).Sum;
}
"Number of chrome processes to create: $numberOfPhysicalCores (# physical cores)";
"Each chrome process requires 1-2GB RAM!"

$directories = @();
foreach ($destinationDir in $args)
{
    $directories += ,$destinationDir;
}
if ($directories.Length -eq 0)
{
    # If TestNameDirectoryFilter is provided, filter directories by that pattern
    if ($TestNameDirectoryFilter) {
        Write-Host "Filtering test directories with pattern: *$TestNameDirectoryFilter*"
        foreach ($destinationDir in Get-ChildItem -Path $workingDir -Directory -Exclude "DemoLoad*" | 
            Where-Object { $_.Name -like "*$TestNameDirectoryFilter*" })
        {
            $directories += ,$destinationDir.Name;
        }
    }
    else {
        # Original behavior - get all directories except DemoLoad*
        foreach ($destinationDir in Get-ChildItem -Path $workingDir -Directory -Exclude "DemoLoad*")
        {
            $directories += ,$destinationDir.Name;
        }
    }
}

if ( $UseDocker -eq $true) { 
    $ContainerPath = "/usr/src/"
    CheckForSeleniumBIModule
    dockerRun
    foreach ($destinationDir in $directories)
    {

        $reportHtmlFile = $(Join-Path (Join-Path $workingDir $destinationDir) $htmlFileName);
        if (Test-Path -path $reportHtmlFile)
        {
            $loopCounter = [int]$instances
            while($loopCounter -gt 0)
            {
            $reportHtmlFile  
            # Run the container
            $containerInstance = $containerName + $loopCounter
            $containerPort = 4444 + $loopCounter
            Write-Host "Starting the container $containerInstance..."
            docker run -d --name $containerInstance -p $containerPort`:4444 $image
            Start-Sleep -Seconds 5
            Write-Host "Copying $reportHtmlFile to $containerInstance`:$containerPath..."
            docker cp $reportHtmlFile $containerInstance`:$containerPath

            # Open the HTML file in the browser inside the container
            Write-Host "Running $htmlFileName in the browser inside $containerInstance..."
            $remoteFile = "file://$containerPath/$htmlFileName"
            $remoteFile
            $remoteAddress = "http://localhost`:$containerPort/wd/hub"
            $capabilities = @{browserName='chrome'}
            #docker exec $containerInstance google-chrome --headless --disable-gpu --no-sandbox $remoteFile & 
            $driver = Start-SeRemote -RemoteAddress $remoteAddress -DesiredCapabilities $capabilities
            $driver.Navigate().GoToUrl($remoteFile)
            start-sleep -Seconds 30
            New-SeScreenshot -Driver $driver -Path "C:\Temp\screenshot$containerPort.png"
            --$loopCounter
            
            }
        }
    }

    "Press enter when load test is complete: "
    pause
    $loopCounter = [int]$instances
    while($loopCounter -gt 0)
    {
        $containerInstance = $containerName + $loopCounter
        $containerPort = 4444 + $loopCounter
        Write-Host "Stopping the container $containerInstance..."
        docker stop $containerInstance
        docker rm -f $containerInstance
        $loopCounter--
    }

}
if ($UseDocker -eq $false) {
 $driverList = @()
    foreach ($destinationDir in $directories)
    {
        $reportHtmlFile = $(Join-Path (Join-Path $workingDir $destinationDir) $htmlFileName);
        if (Test-Path -path $reportHtmlFile)
        {
            $loopCounter = [int]$instances
            while($loopCounter -gt 0)
            {
            $reportHtmlFile  
            $driver = Start-SeChrome
            $driver.Navigate().GoToUrl($reportHtmlFile)
            $driverList += $driver
                 --$loopCounter
                sleep -Seconds 1
            }
        }
    }


"Press enter when load test is complete: "
pause
$i = 0 
$allRefreshTimes = @()

# Collect data from all instances
foreach ($driver in $driverList)
{
    $i++
    $e = Find-SeElement -driver $driver -id LoadReportCounter
    $resultText = $e.Text
    Write-Host "Browser instance $i raw result: $resultText"
    
    # Extract the average refresh time if available
    if ($resultText -match '(\d+\.?\d*) seconds average refresh time') {
        $avgRefreshTime = [double]$matches[1]
        $allRefreshTimes += $avgRefreshTime
    }
    
    $driver.Quit()
}

# Calculate statistics if we have data
if ($allRefreshTimes.Count -gt 0) {
    # Sort for calculating median and getting min/max
    $sortedTimes = $allRefreshTimes | Sort-Object
    
    # Calculate statistics
    $minTime = $sortedTimes[0]
    $maxTime = $sortedTimes[-1]
    $avgTime = ($sortedTimes | Measure-Object -Average).Average
    
    # Calculate median
    if ($sortedTimes.Count % 2 -eq 0) {
        # Even number of items
        $medianIndex = $sortedTimes.Count / 2
        $median = ($sortedTimes[$medianIndex-1] + $sortedTimes[$medianIndex]) / 2
    } else {
        # Odd number of items
        $median = $sortedTimes[($sortedTimes.Count - 1) / 2]
    }
    
    # Display statistics
    Write-Host "`nRefresh Time Statistics (seconds):" -ForegroundColor Cyan
    Write-Host "Minimum:  $($minTime.ToString("0.000"))" -ForegroundColor Green
    Write-Host "Maximum:  $($maxTime.ToString("0.000"))" -ForegroundColor Green
    Write-Host "Average:  $($avgTime.ToString("0.000"))" -ForegroundColor Green
    Write-Host "Median:   $($median.ToString("0.000"))" -ForegroundColor Green
    Write-Host "Samples:  $($allRefreshTimes.Count)" -ForegroundColor Green
}
}