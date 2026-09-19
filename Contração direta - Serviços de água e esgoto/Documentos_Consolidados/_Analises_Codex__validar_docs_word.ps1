$ErrorActionPreference = 'Stop'

$base = 'C:\Users\SAMUEL~1\Desktop\PROCES~1\CONTRA~2\DOCUME~1'
$files = @(
    @('DOCGUA~2.DOC', 'DOC_AGUA_00_Roteiro_uma_pagina.pdf'),
    @('DOCGUA~1.DOC', 'DOC_AGUA_06_Atesto_tecnico_rascunho.pdf'),
    @('DOCGUA~4.DOC', 'DOC_AGUA_07_TR_rascunho.pdf')
)

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

try {
    foreach ($file in $files) {
        $input = Join-Path $base $file[0]
        $output = Join-Path $base $file[1]
        $document = $word.Documents.Open($input, $false, $true)
        try {
            $document.ExportAsFixedFormat($output, 17)
        }
        finally {
            $document.Close(0)
        }
    }
}
finally {
    $word.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
