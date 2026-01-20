Add-Type -AssemblyName System.Windows.Forms
Import-Module ActiveDirectory


$OUMap = @{
  "Anna"   = "OU=Anna,OU=Students,OU=HSWW,DC=ivan,DC=gudz"
  "Dima"   = "OU=Dima,OU=Students,OU=HSWW,DC=ivan,DC=gudz"
  "Kostya" = "OU=Kostya,OU=Students,OU=HSWW,DC=ivan,DC=gudz"
  "Vanya"  = "OU=Vanya,OU=Students,OU=HSWW,DC=ivan,DC=gudz"
}

$form = New-Object Windows.Forms.Form
$form.Text = "New AD User"
$form.Width = 360
$form.Height = 260

function AddField($labelText, $top, $isPassword=$false) {
  $lbl = New-Object Windows.Forms.Label
  $lbl.Text = $labelText
  $lbl.Left = 10; $lbl.Top = $top; $lbl.Width = 120
  $tb = New-Object Windows.Forms.TextBox
  $tb.Left = 140; $tb.Top = $top; $tb.Width = 190
  if ($isPassword) { $tb.UseSystemPasswordChar = $true }
  $form.Controls.AddRange(@($lbl,$tb))
  return $tb
}

$tbFirst = AddField "FirstName" 20
$tbLast  = AddField "LastName"  50
$tbUser  = AddField "UserName"  80
$tbPass  = AddField "Password" 110 $true

$lblOU = New-Object Windows.Forms.Label
$lblOU.Text = "Folder (OU)"
$lblOU.Left = 10; $lblOU.Top = 140; $lblOU.Width = 120

$cbOU = New-Object Windows.Forms.ComboBox
$cbOU.Left = 140; $cbOU.Top = 140; $cbOU.Width = 190
$cbOU.DropDownStyle = "DropDownList"
[void]$cbOU.Items.AddRange($OUMap.Keys)

$btn = New-Object Windows.Forms.Button
$btn.Text = "Create"
$btn.Left = 140; $btn.Top = 175; $btn.Width = 80

$btn.Add_Click({
  if (-not ($tbFirst.Text -and $tbLast.Text -and $tbUser.Text -and $tbPass.Text -and $cbOU.SelectedItem)) {
    [Windows.Forms.MessageBox]::Show("Missing field(s).") | Out-Null
    return
  }

  $ouPath = $OUMap[$cbOU.SelectedItem]
  $secure = ConvertTo-SecureString $tbPass.Text -AsPlainText -Force

  try {
    New-ADUser `
      -Name "$($tbFirst.Text) $($tbLast.Text)" `
      -GivenName $tbFirst.Text `
      -Surname $tbLast.Text `
      -SamAccountName $tbUser.Text `
      -UserPrincipalName "$($tbUser.Text)@ivan.gudz" `
      -Path $ouPath `
      -AccountPassword $secure `
      -Enabled $true
    [Windows.Forms.MessageBox]::Show("Created in: $($cbOU.SelectedItem)") | Out-Null
  } catch {
    [Windows.Forms.MessageBox]::Show($_.Exception.Message) | Out-Null
  }
})

$form.Controls.AddRange(@($lblOU,$cbOU,$btn))
[void]$form.ShowDialog()
