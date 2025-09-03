# Test GUI to verify Forms are working

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = "Test GUI"
$form.Size = New-Object System.Drawing.Size(400, 300)
$form.StartPosition = "CenterScreen"

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(50, 50)
$label.Size = New-Object System.Drawing.Size(300, 50)
$label.Text = "If you can see this, Windows Forms is working!"
$form.Controls.Add($label)

$button = New-Object System.Windows.Forms.Button
$button.Location = New-Object System.Drawing.Point(150, 120)
$button.Size = New-Object System.Drawing.Size(100, 30)
$button.Text = "Click Me"
$button.Add_Click({
    [System.Windows.Forms.MessageBox]::Show("Button clicked!", "Success")
})
$form.Controls.Add($button)

$form.ShowDialog()