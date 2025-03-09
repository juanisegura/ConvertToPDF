$file = $args[0]
$ext = [System.IO.Path]::GetExtension($file).ToLower()
$out = [System.IO.Path]::ChangeExtension($file, ".pdf")
if ($ext -in @(".doc", ".docx")) { 
    $app = New-Object -ComObject Word.Application
    $doc = $app.Documents.Open($file)
    $doc.SaveAs([ref]$out, [ref]17)
    $doc.Close()
    $app.Quit()
}
elseif ($ext -in @(".xls", ".xlsx")) { 
    $app = New-Object -ComObject Excel.Application
    $wb = $app.Workbooks.Open($file)
    $wb.ExportAsFixedFormat(0, $out)
    $wb.Close()
    $app.Quit()
}
elseif ($ext -in @(".ppt", ".pptx")) { 
    $app = New-Object -ComObject PowerPoint.Application
    $pres = $app.Presentations.Open($file)
    $pres.SaveAs($out, 32)
    $pres.Close()
    $app.Quit()
}
