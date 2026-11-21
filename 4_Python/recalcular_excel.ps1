# Abre un .xlsx en Excel (oculto), recalcula, guarda y cierra. Solo toca el archivo indicado.
param([Parameter(Mandatory)][string]$Ruta)
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Open($Ruta)
    $xl.CalculateFull()
    $wb.Save(); $wb.Close($false)
    "recalculado: $Ruta"
} finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
