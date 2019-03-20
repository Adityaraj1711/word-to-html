######################################################################
#                         README
#                         ------
# This file is used to convert word files to html
# Function Wrd-HTML is converting word files to html using Microsoft word API
# and in __FILES__ directory storing html files using LibreOffice API.
# then iterative call to a python script with args as location of html file.
# to compare and save changes to the microsoft html file in the 
# 
# Check if input is file or directory
# 
# author - QLS
######################################################################

$Input = $args[0]

#Define Office Formats
$Wrd_Array = '*.docx', '*.doc', '*.odt', '*.rtf', '*.txt', '*.wpd'
$Exl_Array = '*.xlsx', '*.xls', '*.ods', '*.csv'
$Pow_Array = '*.pptx', '*.ppt', '*.odp'
$Pub_Array = '*.pub'
$Vis_Array = '*.vsdx', '*.vsd', '*.vssx', '*.vss'
$Off_Array = $Wrd_Array + $Exl_Array + $Pow_Array + $Pub_Array + $Vis_Array
$ExtChk    = [System.IO.Path]::GetExtension($Input)

#Convert Word to HTML
Function Wrd-HTML($f, $p)
{
    $Wrd     = New-Object -ComObject Word.Application
    $Version = $Wrd.Version
    $Doc     = $Wrd.Documents.Open($f)
    $saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatFilteredHTML");
    #Check Version of Office Installed
    If ($Version -eq '16.0' -Or $Version -eq '15.0') {
        #$Doc.SaveAs($p, 17)
        $Doc.SaveAs($p, [ref]$saveFormat) 
        $Doc.Close($False)
    }
    ElseIf ($Version -eq '14.0') {
        #$Doc.SaveAs([ref] $p,[ref] 17)
        $Doc.SaveAs($p, [ref]$saveFormat)
        $Doc.Close([ref]$False)
    }
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
    $Wrd.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Wrd)
    Remove-Variable Wrd
    
    $dir   = Split-Path -Path $p
    $LibreDir = $dir + '\' + "___FILES___"

    $Location = Get-Location
    $ParentDir = $Location.Path + '\' + "script.py"
    soffice --headless --convert-to html:"XHTML Writer File:UTF8" $f --outdir $LibreDir
    Start-Sleep -s 10
    
    python $ParentDir $p
}


#Check for Word Formats
Function Wrd-Chk($f, $e, $p){
    $f = [string]$f
    For ($i = 0; $i -le $Wrd_Array.Length; $i++) {
        $Temp = [string]$Wrd_Array[$i]
        $Temp = $Temp.TrimStart('*')
        If ($e -eq $Temp) {
            Wrd-HTML $f $p
        }
    }
}


If ($ExtChk -eq '')
{
    $Files = Get-ChildItem -path $Input -include $Off_Array -recurse
    ForEach ($File in $Files) {
        $Path     = [System.IO.Path]::GetDirectoryName($File)
        $Filename = [System.IO.Path]::GetFileNameWithoutExtension($File)
        $Ext      = [System.IO.Path]::GetExtension($File)
        $PDF      = $Path + '\' + $Filename + '.pdf'
        $html     = $Path + '\' + $Filename + '.html'
        Wrd-Chk $File $Ext $html
    }
}
Else
{
    $File     = $Input
    $Path     = [System.IO.Path]::GetDirectoryName($File)
    $Filename = [System.IO.Path]::GetFileNameWithoutExtension($File)
    $Ext      = [System.IO.Path]::GetExtension($File)
    $PDF      = $Path + '\' + $Filename + '.pdf'
    $html     = $Path + '\' + $Filename + '.html'
    Wrd-Chk $File $Ext $html
}

#Cleanup
Remove-Item Function:Wrd-HTML, Function:Wrd-Chk
