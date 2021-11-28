$Giden = "$Env:userprofile\Desktop\PDF_Parse\Giden"
$Gelen = "$Env:userprofile\Desktop\PDF_Parse\Return"
$Working_Area = "$Env:userprofile\Desktop\Parse"
$Alpha = 0

While ($true) {

    $Files = GET-CHILDITEM -include *.pdf -recurse -path $Giden
    if ($Files.Length -eq 0) {
        Start-Sleep 10

        if ($Alpha -eq 1) {
            Clear-Host
            Write-Host "Bekleniyor..."
            $Alpha = 0
        }
        else {
            Clear-Host
            Write-Host "Bekleme devam ediyor..."
            $Alpha = 1
        }

    }
    else {
        
        Clear-Host
        Remove-Item -R "$Working_Area\" -ErrorAction Ignore
        New-Item -Path "$Env:userprofile\Desktop" -Name "Parse" -ItemType "directory" -ErrorAction Ignore

        Foreach ($File in $Files) {
            [string]$F = $File.Name
            [string]$F1 = $Giden + "\" + $F
            [string]$F2 = "$Working_Area\" + $F
            Copy-Item $F1 -Destination $F2
            Remove-Item -R "$Giden\$F" -ErrorAction Ignore
        }

        $Word = NEW-OBJECT –COMOBJECT WORD.APPLICATION  
        $Files = GET-CHILDITEM -include *.pdf -recurse -path "$Working_Area\"

        Foreach ($File in $Files) {
            try {
                write-host "Deneniyor " $File.fullname 
                $Doc = $Word.Documents.Open($File.fullname)
                write-host "Cevirdi   " $File.fullname ". RESULT=" $?
                $Name = ($File.Fullname).replace("pdf", “docx”)
                $Doc.saveas([ref] $Name, [ref] 16)
            }
            catch { 
                write-host "$File hata var!!!"
            } 
            finally {
                $Doc.close()
                if (($i % 200) -eq 0) {
                    [System.GC]::Collect()
                }

            }
        }

        $objWord = New-Object -Com Word.Application

        $ByteKare = "00000111"
        $ByteTab = "00001001"

        $FinalKare = [char]([convert]::ToInt32($ByteKare, 2))
        $FinalKare = [string]$FinalKare

        $FinalTab = [char]([convert]::ToInt32($ByteTab, 2))
        $FinalTab = [string]$FinalTab

        $Files = GET-CHILDITEM -include *.docx -recurse -path "$Working_Area\"

        $CSS = '<meta name="viewport" content="width=device-width, initial-scale=1"><style>html {-ms-text-size-adjust: 100%;-webkit-text-size-adjust: 100%;}html,body {font-family: Verdana, sans-serif;line-height: 1.5;}table {margin-top: 10px;font-size: 100%;border-collapse: collapse;width: 100%;border: 1px solid black;margin-bottom: 10px;}td {padding: 8px;border: 1px solid black;text-align: left;}tr:nth-child(odd) {background-color: #dddddd;}::-webkit-scrollbar {width: 10px;}::-webkit-scrollbar-track {background: #f1f1f1;}::-webkit-scrollbar-thumb {background: #888;}::-webkit-scrollbar-thumb:hover {background: #555;}.btn {border: none;color: white;padding: 14px 28px;font-size: 16px;cursor: pointer;width: 160px;}.info {background-color: #2196F3;}.info:hover {background: #0b7dda;} #snackbarCikti {visibility: hidden;min-width: 250px;margin-left: -125px;background-color: #333;color: #fff;text-align: center;border-radius: 2px;padding: 16px;position: fixed;z-index: 1;left: 50%;bottom: 30px;} #snackbarCikti.show {visibility: visible;-webkit-animation: fadein 0.5s, fadeout 0.5s 2.5s;animation: fadein 0.5s, fadeout 0.5s 2.5s;} @-webkit-keyframes fadein {from {bottom: 0;opacity: 0;}to {bottom: 30px;opacity: 1;}} @keyframes fadein {from {bottom: 0;opacity: 0;}to {bottom: 30px;opacity: 1;}} @-webkit-keyframes fadeout {from {bottom: 30px;opacity: 1;}to {bottom: 0;opacity: 0;}} @keyframes fadeout {from {bottom: 30px;opacity: 1;}to {bottom: 0;opacity: 0;}}</style>'
$JS = '<script type="text/javascript">function Excel_click() {try { var table_html = document.getElementById("Tablo").outerHTML; } catch (err) { var table_html = 0 };if (table_html != 0) {var dt = new Date();var day = dt.getDate();var month = dt.getMonth() + 1;var year = dt.getFullYear();var date_last = day + "." + month + "." + year;var csvString = "ı,ü,ü,ğ,ş,#Hashtag,ä,ö";var universalBOM = "\uFEFF";var data_type = ("href", "data:text/csv; charset=utf-8," + encodeURIComponent(universalBOM + csvString));var css_html = "<style>td {border: 0.5pt solid #c0c0c0} .tRight { text-align:right} .tLeft { text-align:left}th { border: 1px solid black; text-align: center; background-color: #002060; color: white;} </style>";var a = document.createElement("a");a.href = data_type + "," + encodeURIComponent("<html><head>" + css_html + "</head><body>" + table_html + "</body></html>");a.download = "Firma_" + date_last + ".xls";a.click();document.getElementById("snackbarCikti").textContent = "Excel oluşturuldu."} else {document.getElementById("snackbarCikti").textContent = "Excele aktarılabilecek veri bulunamadı."}var x = document.getElementById("snackbarCikti");x.className = "show";setTimeout(function () { x.className = x.className.replace("show", ""); }, 3000);};</script>'

        Foreach ($File in $Files) {
            $Tablo = ""
            $F = $File.name
            $filename = "$Working_Area\" + $F
            $F = $F.replace(".docx", "")

            $objDocument = $objWord.Documents.Open($filename)
            $Counter = 1
            $objDocument.Tables | ForEach-Object {
    
                $UTablo = $objDocument.Tables.Item($Counter)
                $UTabloCols = $UTablo.Columns.Count
                $UTabloRows = $UTablo.Rows.Count
    
                for ($a = 1; $a -le $UTabloRows; $a++) {
                    $Veris = ""
                    $x = 0
                    for ($b = 1; $b -le $UTabloCols; $b++) {
                        $Veri = $UTablo.Cell($a, $b).Range.Text
        
                        if ($x -eq 0) {
                            $Veri = $Veri.replace($FinalKare, "")
                            $Veri = $Veri.replace($FinalTab, " ")
                            $Veris = $Veri.trim()
                            $Veris = '<td>' + $Veris + '</td>'

                            $x = 1
                        }
                        else {
                            $Veri = $Veri.replace($FinalKare, "")
                            $Veri = $Veri.replace($FinalTab, " ")
                            $Veri = $Veri.trim()
                            $Veris = $Veris + '<td>' + $Veri + '</td>'
                        }
        
                    }

                    $Tablo = $Tablo + '<tr>' + $Veris + '</tr>'

                }
   
                $Counter = $Counter + 1
            }

            $Date = Get-Date -format "yyyyMMdd-HH-mm-ss-fff"

            New-Item -Path "$Gelen" -Name $Date -ItemType "directory"

            [string]$CiktiLink = "$Gelen\$Date\ParseEdilen.html"

            $Tablo = '<html><head>' + $CSS + '</head><body><p>' + $F + ' isimli dosya ' + $Date + ' tarihinde parse edilmiştir.</p><button id="Excel" onclick="Excel_click()" title="Excel" class="btn info">Excel</button><div id="snackbarCikti"></div><table id = "Tablo">' + $Tablo + '</table>' + $JS + '</body></html>'
            $Tablo > $CiktiLink
            Copy-Item "$Working_Area\$F.pdf" -Destination "$Gelen\$Date\$F.pdf"
            Clear-Host
        }

        $objDocument.Close()
        $objWord.Quit()

        $wd = [Runtime.Interopservices.Marshal]::GetActiveObject('Word.Application')
        $wd.Documents | ForEach-Object { $_.Close() }
        Get-Process | Where-Object { $_.ProcessName -eq "WINWORD" } | Stop-Process

        Remove-Item -R "$Working_Area\" –recurse -ErrorAction Ignore

        Clear-Host

        if (($i % 200) -eq 0) {
            [System.GC]::Collect()
        }

    }

}