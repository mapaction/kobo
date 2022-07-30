<#
        .SYNOPSIS
        When supplied with a standard Administrative Divisions Excel file, produces a cascading menu options for use in a Kobo survey. 

        .DESCRIPTION
        Produces Excel files in the form required by Kobo: https://support.kobotoolbox.org/cascading_select.html

        .PARAMETER inputAdminBoundariesTabulardataXlsx
        Specifies the full path to the input Excel file.

        These can be downloaded from HDX, e.g.:
           https://data.humdata.org/dataset?ext_administrative_divisions=1&q=&sort=if(gt(last_modified%2Creview_date)%2Clast_modified%2Creview_date)%20desc&ext_page_size=25

        Typically, they are named:

           <three-letter country code>_adminboundaries_tabulardata.xlsx

        For example, for Burundi:
           bdi_adminboundaries_tabulardata.xlsx

        .PARAMETER outputXlsx
        Full path to output Excel file

        .PARAMETER overwrite
        
        Defaults to $false

        If supplied as $true, then will over-write existing output file if it exists.

        If supplied as $false and the path to the output file does not exist, then the script will raise an exception and exit.
        
        .PARAMETER traceFlag
        Defaults to $false

        If supplied as $true, then will produce output as trace.

        .EXAMPLE
        PS> .\generateCascadingQuestionsFromAdminBoundaries.ps1 `
            -inputAdminBoundariesTabulardataXlsx C:\Users\steve\source\repos\kobo\utility\generateCascadingQuestionsFromAdminBoundaries\bdi_adminboundaries_tabulardata.xlsx `
            -outputXlsx C:\Users\steve\source\repos\kobo\utility\generateCascadingQuestionsFromAdminBoundaries\bdi_cascaded_selections.xlsx `
            -overwrite $true `
            -traceFlag $false

        .TESTING
        Tested against:
           afg_adminboundaries_tabulardata.xlsx
           bdi_adminboundaries_tabulardata.xlsx
           hun_adminboundaries_tabulardata.xlsx
           mdg_gazetteer_20181031.xlsx
    #>

param
   (
      [Parameter(Mandatory=$true)][String]$inputAdminBoundariesTabulardataXlsx,
      [Parameter(Mandatory=$true)][String]$outputXlsx,
      [Boolean]$overwrite=$false,
      [Boolean]$traceFlag=$false
   )

function Trace
{
    param([String]$message)
    if ($traceFlag)
    {
        Write-Host Trace: $message
    }
}

function SheetNaming
{
    # Most common elements in an array
    param($sheetNamesArray)

        $sheetNamesMinusNumbers = @()
        foreach ($sheetName in $sheetNamesArray)
        {
            $sheetNamesMinusNumbers += , $sheetName.Name -replace '\d+',''
            # Write-Host ($sheetName.Name -replace '\d+','')
        }
        $naming = $sheetNamesMinusNumbers | group | sort count -desc | select -ExpandProperty Name -First 1

        # Write-Host ($naming)
        return $naming
}

function columnHeadingParts
{
    # Most common elements in an array
    param($workSheet, $smallestAdminLevel)

        $columns = $workSheet.UsedRange.Columns.Count
        $headingRowNumber = 1
        $suffixes = @()
        $parts= @()
        $pCodePrefix = ""
        $pCodeSuffix = ""
        $namePrefix = ""
        $nameSuffix = ""
        for ($j=1; $j -le $columns; $j++)
        {
            if ($workSheet.Cells.Item($headingRowNumber,$j).text -like ("*" + $smallestAdminLevel.ToString() + "*"))
            {
                $parts = ($workSheet.Cells.Item($headingRowNumber, $j).text).Split($smallestAdminLevel.ToString())
                if (($workSheet.Cells.Item($headingRowNumber, $j).text).ToUpper().Contains("PCODE"))
                {
                    $pCodePrefix = $parts[0]
                    $pCodeSuffix = $parts[1]
                }
                else
                {
                    if (($workSheet.Cells.Item($headingRowNumber, $j).text).ToUpper().Contains("ALT"))
                    {
                    }
                    elseif (($workSheet.Cells.Item($headingRowNumber, $j).text).ToUpper().Contains("REF"))
                    {
                    }
                    else
                    {
                        $namePrefix = $parts[0]
                        $nameSuffix = ($parts[1] -replace ".{3}$")
                    }
                }                
            }
        }
        return $pCodePrefix, $pCodeSuffix, $namePrefix, $nameSuffix
}

if( -Not ($inputAdminBoundariesTabulardataXlsx | Test-Path) ){
   throw "File or folder does not exist"
}

if ($outputXlsx | Test-Path)
{
   if ($overwrite -eq $true)
   {
      Remove-Item $outputXlsx -Force
   }
   else
   {
      throw ("File : " + $outputXlsx + " already exists")
   }
}

Trace $inputAdminBoundariesTabulardataXlsx

$excelInput = new-object -comobject Excel.Application
$workbook = $excelInput.workbooks.Open($inputAdminBoundariesTabulardataXlsx)
$sheets = $workbook.sheets

$sheetArray = @()
# Get naming convention for the Worksheets, e.g.
#
# mdg_gazetteer_20181031.xlsx has 'mdg_pop_adm0', 'mdg_pop_adm1', 'mdg_pop_adm2', 'mdg_pop_adm4', 'mdg_pop_adm4'
# hun_adminboundaries_tabulardata.xlsx 
# bdi_adminboundaries_tabulardata.xlsx
#   have:
#      'Admin0', 'Admin1', 'Admin2'  

$sheetNamingPrefix = SheetNaming -sheetNamesArray $sheets

# We won't know which is the smallest admin level in the spreadsheet
$sheets | % { ` 
   if ($_.name.StartsWith($sheetNamingPrefix))
   {
      $sheetArray += $_.name
   }
}
$smallestAdminLevelWorksheet = ($sheetArray | sort)[-1]
$smallestAdminLevel = $smallestAdminLevelWorksheet  -replace "[^0-9]" , ''

$workSheet = $workbook.Sheets.Item($smallestAdminLevelWorksheet)

# Open the smallest admin level worksheet
Trace ("Worksheet Name : " + $workSheet.Name + " contains " + $workSheet.UsedRange.Rows.Count.ToString() + " rows.")

$pCodePrefix, $pCodeSuffix, $namePrefix, $nameSuffix = columnHeadingParts -workSheet $workSheet -smallestAdminLevel $smallestAdminLevel

$excelOutput = new-object -comobject Excel.Application
$workbookOutput = $excelOutput.Workbooks.Add()
$workSheetOutput = $workbookOutput.Worksheets.Item(1)
$workSheetOutput.Cells.Item(1,1) = "list_name"
$workSheetOutput.Cells.Item(1,2) = "name"
$workSheetOutput.Cells.Item(1,3) = "label"
$columnNumber = 4
$columns = $workSheet.UsedRange.Columns.Count
$headingRowNumber = 1
$nameColumnNumbers = @()
$pCodeColumnNumbers = @()
$recordArray = @()
$recordCollection = @()
$mostCommonNameColumnNameSuffix = ""
$suffixes = @()

for ($level=0; $level -le $smallestAdminLevel; $level++)
{
#    Write-Host ("TSTSV : " + ($namePrefix + $level + $nameSuffix + '_[a-z][a-z]'))
    for ($j=1; $j -le $columns; $j++)
    {
        #if ($workSheet.Cells.Item($headingRowNumber,$j).text -match ($valuePrefix + $level + 'Name_[a-z][a-z]'))
        if ($workSheet.Cells.Item($headingRowNumber,$j).text -imatch ($namePrefix + $level + $nameSuffix + '_[a-z][a-z]'))
        {
#    Write-Host ("YAYTSTSV : " + ($namePrefix + $level + $nameSuffix + '_[a-z][a-z]'))
            $length = $workSheet.Cells.Item($headingRowNumber,$j).text.length
            $lastTwoCharacters = ($workSheet.Cells.Item($headingRowNumber,$j).text).substring($length -2)
            $suffixes += , $lastTwoCharacters 
        }
    }
}
$mostCommonNameColumnNameSuffix = ($suffixes | group | sort count -desc | select -f 1).Name

# For each Administrative Level, get the AdminName and PCode column numbers 
for ($level=0; $level -le $smallestAdminLevel; $level++)
{
    for ($j=1; $j -le $columns; $j++)
    {
        if ($workSheet.Cells.Item($headingRowNumber,$j).text -eq ($namePrefix + $level + $nameSuffix + "_" + $lastTwoCharacters))
        {
            $nameColumnNumbers += $j
        }
        elseif ($workSheet.Cells.Item($headingRowNumber,$j).text -eq ($pCodePrefix + $level + $pCodeSuffix))
        {
            $pCodeColumnNumbers += $j
        }
    }
    if ($level -lt $smallestAdminLevel)
    {
        $workSheetOutput.Cells.Item(1,$columnNumber) = $namePrefix + $level
    }
    Trace("Admin Level " + $level.ToString() + " : Name  column number is column " + $nameColumnNumbers[$level].ToString())
    Trace("Admin Level " + $level.ToString() + " : PCode column number is column " + $pCodeColumnNumbers[$level].ToString())

    $columnNumber = $columnNumber + 1
    $recordArray += , @()
    $recordCollection += , @() 
}

# Start at second row - after the header row
for ($r=2; $r -le $workSheet.UsedRange.Rows.Count; $r++)
{
   for ($c = 0; $c -lt $nameColumnNumbers.Count; $c++)
   {
      if ($recordArray[$c] -notcontains $workSheet.Cells.Item($r,$pCodeColumnNumbers[$c]).text)
      {
         $recordArray[$c] += $workSheet.Cells.Item($r,$pCodeColumnNumbers[$c]).text
         if ($workSheet.Cells.Item($r,$pCodeColumnNumbers[$c]).text.Length -gt 0)
         {
            $recordCollection[$c] += New-Object PSObject -Property @{
                                                                       ListName = ($namePrefix + $c)
                                                                       Name = $workSheet.Cells.Item($r,$pCodeColumnNumbers[$c]).text
                                                                       Label = $workSheet.Cells.Item($r,$nameColumnNumbers[$c]).text
                                                                       PreviousAdminValue = $workSheet.Cells.Item($r,$pCodeColumnNumbers[($c-1)]).text
                                                                    }
            Trace("Level: " + $c.ToString() + ": Cell (" + $r + ", " + $pCodeColumnNumbers[($c)] + ") = " + $workSheet.Cells.Item($r,$pCodeColumnNumbers[($c)]).text + " Level: " + ($c-1).ToString() + ": Cell (" + $r + ", " + $pCodeColumnNumbers[($c-1)] + ") = " + $workSheet.Cells.Item($r,$pCodeColumnNumbers[($c-1)]).text)
         } 
      }   
   } 
}

$completedSuccessfully=$true
Trace("Record counts for the different Administrative Levels:")
for ($rc = 0; $rc -lt $recordCollection.Count; $rc++)
{
   Trace("Level " + $rc + " has " + $recordCollection[$rc].Count + " records")
   if ($recordCollection[$rc].Count -eq 0)
   {
      $completedSuccessfully = $false
   }
}

$previousCount=2 # Header
for ($level=0; $level -le $smallestAdminLevel; $level++)
{
   for ($row = 0; $row -lt $recordArray[$level].Count; $row++)
   {
      $workSheetOutput.Cells.Item(($previousCount+$row),1) = $recordCollection[$level][$row].ListName
      $workSheetOutput.Cells.Item(($previousCount+$row),2) = $recordCollection[$level][$row].Name
      $workSheetOutput.Cells.Item(($previousCount+$row),3) = $recordCollection[$level][$row].Label
      if ($level -gt 0)
      {
         $workSheetOutput.Cells.Item(($previousCount+$row),(3 + $level)) = $recordCollection[$level][$row].PreviousAdminValue
      }
   }
   $previousCount = $previousCount+$recordCollection[$level].Count + 1
}

if ($completedSuccessfully -eq $true)
{
    Write-Host("Output written to " + $outputXlsx)
}
# Close files
$excelInput.Quit()
$releaseResult = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelInput)
Remove-Variable excelInput

$workbookOutput.SaveAs($outputXlsx) 
$excelOutput.Quit()
$releaseResult = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelOutput)
Remove-Variable excelOutput
