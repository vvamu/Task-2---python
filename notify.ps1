Add-Type -AssemblyName System.Windows.Forms

$Form = New-Object system.Windows.Forms.Form
$Form.Text = "Notification"
$Form.Width = 200
$Form.Height = 100

$Label = New-Object System.Windows.Forms.Label
$Label.Text = "Program execution completed."
$Label.AutoSize = $true
$Label.Left = 50
$Label.Top = 20

$Form.Controls.Add($Label)
$Form.Add_Shown({$Form.Activate()})
$Form.ShowDialog()