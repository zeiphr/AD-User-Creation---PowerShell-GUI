#########THE FORM############################################
#                                                           #
  Add-Type -AssemblyName System.Windows.Forms               #
  [System.Windows.Forms.Application]::EnableVisualStyles()  #
#                                                           #
#############################################################


## Get Credentails for Mailbox creation ! ADD DOMAIN FOR FASTER USER CREDENTIAL INPUT##

$UserCredential = Get-Credential "Domain\"


## Please wait - Main form ##

Add-Type -AssemblyName System.Windows.Forms 
$PleaseWait                      = New-Object system.Windows.Forms.Form
$PleaseWait_L                    = New-Object System.Windows.Forms.label #Needs to be below '$PleaseWait = New-Object system.Windows.Forms.Form' 
$PleaseWait.ClientSize           = '715,189'
$PleaseWait.Text                 = "Loading..."
$PleaseWait.Controls.Add($PleaseWait_L)
$PleaseWait.Visible              = $True
$PleaseWait.Icon                 = "\\SMB PATH TO ICON"


## Please Wait form label ##


$PleaseWait_L.width              = 50
$PleaseWait_L.height             = 10
$PleaseWait_L.location           = New-Object System.Drawing.Point(59,79)
$PleaseWait_L.Text               = "Indexing AD data for auto populated fields - Please Wait...."
$PleaseWait_L.AutoSize           = $True
$PleaseWait_L.Font               = 'Microsoft Sans Serif,16'


## Please Wait Sleep - Allows form to fully load ##

Sleep -Seconds 1


## The Main AD User creation box ##

$MainForm                        = New-Object system.Windows.Forms.Form
$MainForm.ClientSize             = '1440,605'
$MainForm.text                   = "Zeiphr - AD User Creation Beta v0.6.27"
$MainForm.TopMost                = $false
$MainForm.Icon                   = "\\SMB PATH TO ICON"


## General Tab label ##

$General                         = New-Object system.Windows.Forms.Label
$General.text                    = "General Tab"
$General.AutoSize                = $true
$General.width                   = 25
$General.height                  = 10
$General.location                = New-Object System.Drawing.Point(35,29)
$General.Font                    = [System.Drawing.Font]::new("Microsoft Sans Serif", 12, [System.Drawing.FontStyle]::Bold)


## Telephones / Account Tab label ##

$TelAcc                          = New-Object system.Windows.Forms.Label
$TelAcc.text                     = "Telephone / Account"
$TelAcc.AutoSize                 = $true
$TelAcc.width                    = 25
$TelAcc.height                   = 10
$TelAcc.location                 = New-Object System.Drawing.Point(500,29)
$TelAcc.Font                     = [System.Drawing.Font]::new("Microsoft Sans Serif", 12, [System.Drawing.FontStyle]::Bold)


## Organization ##

$Org                             = New-Object system.Windows.Forms.Label
$Org.text                        = "Organization"
$Org.AutoSize                    = $true
$Org.width                       = 25
$Org.height                      = 10
$Org.location                    = New-Object System.Drawing.Point(750,29)
$Org.Font                        = [System.Drawing.Font]::new("Microsoft Sans Serif", 12, [System.Drawing.FontStyle]::Bold)


## Address ##

$Address                         = New-Object system.Windows.Forms.Label
$Address.text                    = "Address"
$Address.AutoSize                = $true
$Address.width                   = 25
$Address.height                  = 10
$Address.location                = New-Object System.Drawing.Point(1000,29)
$Address.Font                    = [System.Drawing.Font]::new("Microsoft Sans Serif", 12, [System.Drawing.FontStyle]::Bold)


## GivenName attribute in AD ##

$GivenName                       = New-Object system.Windows.Forms.Textbox
$GivenName.multiline             = $false
$GivenName.width                 = 179
$GivenName.height                = 20
$GivenName.location              = New-Object System.Drawing.Point(35,99)
$GivenName.Font                  = 'Microsoft Sans Serif,10'


## Givenname label - (First Name) ##

$GivenName_L                     = New-Object system.Windows.Forms.Label
$GivenName_L.text                = "First Name - No spaces allowed"
$GivenName_L.AutoSize            = $true
$GivenName_L.width               = 25
$GivenName_L.height              = 10
$GivenName_L.location            = New-Object System.Drawing.Point(35,79)
$GivenName_L.Font                = 'Microsoft Sans Serif,10'


## Surname attribute in AD ##

$Surname                         = New-Object system.Windows.Forms.TextBox
$Surname.multiline               = $false
$Surname.width                   = 179
$Surname.height                  = 20
$Surname.location                = New-Object System.Drawing.Point(252,99)
$Surname.Font                    = 'Microsoft Sans Serif,10'

## Surname label - (Last Name) ##

$Surname_L                       = New-Object system.Windows.Forms.Label
$Surname_L.text                  = "Last Name - No spaces allowed"
$Surname_L.AutoSize              = $true
$Surname_L.width                 = 25
$Surname_L.height                = 10
$Surname_L.location              = New-Object System.Drawing.Point(252,79)
$Surname_L.Font                  = 'Microsoft Sans Serif,10'


## Description attribute in AD ##

$Description                     = New-Object system.Windows.Forms.TextBox
$Description.multiline           = $false
$Description.width               = 179
$Description.height              = 20
$Description.location            = New-Object System.Drawing.Point(35,167)
$Description.Font                = 'Microsoft Sans Serif,10'


## Description label - (Job Description / Title ##

$Description_L                   = New-Object system.Windows.Forms.Label
$Description_L.text              = "Job Description / Title"
$Description_L.AutoSize          = $true
$Description_L.width             = 25
$Description_L.height            = 10
$Description_L.location          = New-Object System.Drawing.Point(35,146)
$Description_L.Font              = 'Microsoft Sans Serif,10'


## TelephoneNumber attribute in AD ##

$TelephoneNumber                 = New-Object system.Windows.Forms.TextBox
$TelephoneNumber.multiline       = $false
$TelephoneNumber.width           = 179
$TelephoneNumber.height          = 20
$TelephoneNumber.location        = New-Object System.Drawing.Point(252,167)
$TelephoneNumber.Font            = 'Microsoft Sans Serif,10'


## TelephoneNumber label - (DDI - If Applicable) ##

$TelephoneNumber_L               = New-Object system.Windows.Forms.Label
$TelephoneNumber_L.text          = "DDI - If Applicable"
$TelephoneNumber_L.AutoSize      = $true
$TelephoneNumber_L.width         = 25
$TelephoneNumber_L.height        = 10
$TelephoneNumber_L.location      = New-Object System.Drawing.Point(252,146)
$TelephoneNumber_L.Font          = 'Microsoft Sans Serif,10'


## Mobile/MobileNumber attributes in AD##    ##Note: Two fields exist in AD with a mobile number ##

$Mobile                          = New-Object system.Windows.Forms.TextBox
$Mobile.multiline                = $false
$Mobile.width                    = 179
$Mobile.height                   = 20
$Mobile.location                 = New-Object System.Drawing.Point(500,99)
$Mobile.Font                     = 'Microsoft Sans Serif,10'


## Mobile/MobileNumber label - (Mobile Number / VPN Auth) ##

$Mobile_L                        = New-Object system.Windows.Forms.Label
$Mobile_L.text                   = "Mobile Number / VPN Auth"
$Mobile_L.AutoSize               = $true
$Mobile_L.width                  = 25
$Mobile_L.height                 = 10
$Mobile_L.location               = New-Object System.Drawing.Point(500,79)
$Mobile_L.Font                   = 'Microsoft Sans Serif,10'


## Department attribte in AD ##

#OU Search veriables##########################################################################################################################################################
#                                                                                                                                                                            #
  $OU                            = 'OU=Users OU,DC=Domain,DC=co,DC=nz'                                                                                                       #                 
  $AllDepartments                = Get-ADUser -SearchBase $OU -Filter * -SearchScope Subtree -Properties Department | Select-Object Department | sort department -Unique     #
  $FilteredDepartments           = $AllDepartments| Where-Object{$AllDepartments.Department} | Sort-Object                                                                   #
#                                                                                                                                                                            #
##############################################################################################################################################################################

$Department                      = New-object System.Windows.Forms.ComboBox
$Department.text                 = "Start typing Department.."
$Department.width                = 180
$Department.height               = 20
$Department.location             = New-Object System.Drawing.Point(750,99)
$Department.Font                 = 'Microsoft Sans Serif,10'
$Department.DropDownWidth        = 250
$Department.AutoCompleteMode     = 3
$Department.AutoCompleteSource   = 256
 
#Placing Departments logic into form###################
#                                                     #
    foreach($AllDepartments in $FilteredDepartments)  #
  {                                                   #
  $Department.Items.Add(($AllDepartments.department)) #
   }                                                  #
#######################################################


## Department Label - (Department) ##

$Department_L                    = New-Object system.Windows.Forms.Label
$Department_L.text               = "Department"
$Department_L.AutoSize           = $true
$Department_L.width              = 25
$Department_L.height             = 10
$Department_L.location           = New-Object System.Drawing.Point(750,79)
$Department_L.Font               = 'Microsoft Sans Serif,10'




## Users Password output ##     ### Note: This has to be displayed last in order for the 'tab' functionality to work on the form ###

$Password_L                      = New-Object system.Windows.Forms.Label
$Password_L.text                 = "Password (Recommended)"
$Password_L.AutoSize             = $true
$Password_L.width                = 25
$Password_L.height               = 10
$Password_L.location             = New-Object System.Drawing.Point(500,146)
$Password_L.Font                 = 'Microsoft Sans Serif,10'
$Password_L.ForeColor            = 'RED'


function Get-RandomCharacters($length, $characters) {
    $random = 1..$length | ForEach-Object { Get-Random -Maximum $characters.length }
    $private:ofs=""
    return [String]$characters[$random]
}
 
function Scramble-String([string]$inputString){     
    $characterArray = $inputString.ToCharArray()   
    $scrambledStringArray = $characterArray | Get-Random -Count $characterArray.Length     
    $outputString = -join $scrambledStringArray
    return $outputString 
}
 
$password                        = Get-RandomCharacters -length 5 -characters 'abcdefghiklmnoprstuvwxyz'
$password                       += Get-RandomCharacters -length 2 -characters 'ABCDEFGHKLMNOPRSTUVWXYZ'
$password                       += Get-RandomCharacters -length 1 -characters '1234567890'
$password                       += Get-RandomCharacters -length 2 -characters '!"$%&/()=?}][{@#*+'
$password                        = Scramble-String $password 

## Users Password Displayed ##

$PasswordSecureString            = New-Object system.Windows.Forms.TextBox
$PasswordSecureString.Text       = "$password"
$PasswordSecureString.ReadOnly   = $True
$PasswordSecureString.multiline  = $false
$PasswordSecureString.width      = 179
$PasswordSecureString.height     = 20
$PasswordSecureString.location   = New-Object System.Drawing.Point(500,167)
$PasswordSecureString.Font       = 'Microsoft Sans Serif,10'

$passwordstring = $Password


## Manager attribute in AD ##

#OU Search veriables############################################################################################################################
#                                                                                                                                              #
$OUs                                = 'OU=Users OU,DC=Domain,DC=co,DC=nz','OU=2nd Users OU,DC=Domain,DC=co,DC=nz'                              #
$AllManagers                        = $OUs| foreach { Get-ADUser -SearchBase $_ -Filter * -SearchScope Subtree -Properties SamAccountName} `
| Where-Object{$_.name -notlike "*SFB*"} `
| Where-Object{$_.name -notlike "*1*"} `
| Where-Object{$_.name -notlike "*2*"} `
| Where-Object{$_.name -notlike "*3*"} `
| Where-Object{$_.name -notlike "*4*"} `
| Where-Object{$_.name -notlike "*5*"} `
| Where-Object{$_.name -notlike "*6*"} `
| Where-Object{$_.name -notlike "*7*"} `
| Where-Object{$_.name -notlike "*8*"} `
| Where-Object{$_.name -notlike "*9*"} `
| Where-Object{$_.name -notlike "*svc*"} `
| Where-Object{$_.name -notlike "*Common names you don't want included*"} `
| Select-Object SamAccountName | sort SamAccountName -Unique                                                                           
                                                                                                                                                #
  $Filteredmanagers              = $AllManagers| Where-Object{$AllManagers.SamAccountName} | Sort SamAccountName -Descending                    #
#                                                                                                                                               #
#################################################################################################################################################


$Manager                         = New-object System.Windows.Forms.ComboBox
$Manager.text                    = "Start typing name.."
$Manager.width                   = 180
$Manager.height                  = 20
$Manager.location                = New-Object System.Drawing.Point(750,166)
$Manager.Font                    = 'Microsoft Sans Serif,10'
$Manager.DropDownWidth           = 250
$Manager.AutoCompleteMode        = 3
$Manager.AutoCompleteSource      = 256
 

#Placing SamAccountName logic into form##################
#                                                       #
    foreach($Allmanagers in $Filteredmanagers)          #
  {                                                     #
  $Manager.Items.Add(($AllManagers.SamAccountName))     #
   }                                                    #
#                                                       #
#########################################################


## Manager label - (Manager) ##


$Manager_L                       = New-Object system.Windows.Forms.Label
$Manager_L.text                  = "Manager"
$Manager_L.AutoSize              = $true
$Manager_L.width                 = 25
$Manager_L.height                = 10
$Manager_L.location              = New-Object System.Drawing.Point(750,146)
$Manager_L.Font                  = 'Microsoft Sans Serif,10'


## StreetAddress attribute in AD ##

$StreetAddress                   = New-Object system.Windows.Forms.Textbox
$StreetAddress.multiline         = $false
$StreetAddress.width             = 179
$StreetAddress.height            = 20
$StreetAddress.location          = New-Object System.Drawing.Point(1000,99)
$StreetAddress.Font              = 'Microsoft Sans Serif,10'


## StreetAddress label - (Street Address) ##

$StreetAddress_L                 = New-Object system.Windows.Forms.Label
$StreetAddress_L.text            = "Street Address"
$StreetAddress_L.AutoSize        = $true
$StreetAddress_L.width           = 25
$StreetAddress_L.height          = 10
$StreetAddress_L.location        = New-Object System.Drawing.Point(1000,79)
$StreetAddress_L.Font            = 'Microsoft Sans Serif,10'


## City attribute in AD ##

$City                            = New-Object system.Windows.Forms.Textbox
$City.multiline                  = $false
$City.width                      = 179
$City.height                     = 20
$City.location                   = New-Object System.Drawing.Point(1217,99)
$City.Font                       = 'Microsoft Sans Serif,10'


## City label - (First Name) ##

$City_L                          = New-Object system.Windows.Forms.Label
$City_L.text                     = "City"
$City_L.AutoSize                 = $true
$City_L.width                    = 25
$City_L.height                   = 10
$City_L.location                 = New-Object System.Drawing.Point(1217,79)
$City_L.Font                     = 'Microsoft Sans Serif,10'


## State attribute in AD ##

$State                           = New-Object system.Windows.Forms.Textbox
$State.multiline                 = $false
$State.width                     = 179
$State.height                    = 20
$State.location                  = New-Object System.Drawing.Point(1000,167)
$State.Font                      = 'Microsoft Sans Serif,10'


## State label - (First Name) ##

$State_L                         = New-Object system.Windows.Forms.Label
$State_L.text                    = "State"
$State_L.AutoSize                = $true
$State_L.width                   = 25
$State_L.height                  = 10
$State_L.location                = New-Object System.Drawing.Point(1000,146)
$State_L.Font                    = 'Microsoft Sans Serif,10'


## PostCode attribute in AD ##

$PostCode                        = New-Object system.Windows.Forms.Textbox
$PostCode.multiline              = $false
$PostCode.width                  = 179
$PostCode.height                 = 20
$PostCode.location               = New-Object System.Drawing.Point(1217,167)
$PostCode.Font                   = 'Microsoft Sans Serif,10'


## PostCode label - (First Name) ##

$PostCode_L                      = New-Object system.Windows.Forms.Label
$PostCode_L.text                 = "Postcode"
$PostCode_L.AutoSize             = $true
$PostCode_L.width                = 25
$PostCode_L.height               = 10
$PostCode_L.location             = New-Object System.Drawing.Point(1217,146)
$PostCode_L.Font                 = 'Microsoft Sans Serif,10'


## Office attribute in AD ##

#OU Search veriables###############################################################################################################################################################################################################################
#                                                                                                                                                                                                                                                 #
  $OU                            = 'OU=Users OU,DC=Domain,DC=co,DC=nz'                                                                                                                                                                            #
  $OUDescription                 = Get-ADOrganizationalUnit -SearchBase $OU -Filter * -SearchScope subtree | Where-Object{$_.name -notlike "*Computer*"} | Where-Object{$_.name -notlike "User*"} | Select-Object name                            #
  $OfficeLocations               = Get-ADOrganizationalUnit -SearchBase $OU -Filter * -SearchScope Subtree -Properties description | Where-Object{$_.description -like "*Users"} | Select-Object description,DistinguishedName | Sort-Object      #
  $FilteredLocations             = $OfficeLocations | Where-Object{$OfficeLocations.description} | Sort-Object                                                                                                                                    #
#                                                                                                                                                                                                                                                 #
###################################################################################################################################################################################################################################################

$Office                          = New-object System.Windows.Forms.ComboBox
$Office.text                     = "Start typing office location.."
$Office.width                    = 850
$Office.height                   = 46
$Office.location                 = New-Object System.Drawing.Point(35,293)
$Office.Font                     = 'Microsoft Sans Serif,10'
$Office.DropDownWidth            = 850
$Office.AutoCompleteMode         = 3
$Office.AutoCompleteSource       = 256



#Placing SamAccountName into form######################################################################
#                                                                                                     #
foreach($OfficeLocation in $FilteredLocations)                                                        #
{                                                                                                     #
 $Office.Items.add("$($OfficeLocation.description) $(":") $($Officelocation.DistinguishedName)")      #
 }                                                                                                    #
#                                                                                                     #
#######################################################################################################


 ## Office Location label - (Office Loaction) ##

$Office_L                        = New-Object system.Windows.Forms.Label
$Office_L.text                   = "Office location"
$Office_L.AutoSize               = $true
$Office_L.width                  = 25
$Office_L.height                 = 10
$Office_L.location               = New-Object System.Drawing.Point(35,273)
$Office_L.Font                   = 'Microsoft Sans Serif,10'


## Mailbox Note Label - (Please note:) ##

$MailNote_L                      = New-Object system.Windows.Forms.Label
$MailNote_L.text                 = "Please note: Both On-Prem and Office 365 mailboxes will be created with user "
$MailNote_L.AutoSize             = $true
$MailNote_L.width                = 25
$MailNote_L.height               = 10
$MailNote_L.location             = New-Object System.Drawing.Point(940,250)
$MailNote_L.Font                 = 'Microsoft Sans Serif,10'

## Close Please Wait form - (Top of script) ##

$PleaseWait.Close()

## Create user & Close button + Label in one ##

$Button                          = New-Object -TypeName System.Windows.Forms.Button
$Button.text                     = "Create User and close"
$Button.width                    = 220
$Button.height                   = 55
$Button.location                 = New-Object System.Drawing.Point(940,275)
$Button.Font                     = 'Microsoft Sans Serif,12'
$Button.add_Click({

## Takes Office OU + Display name - Strips out Display name and leading space ##

## This bit was added in due to a dynamic DG that we had. Therefore when the OU = a specific OU, it would replace the Office feild with $path 2 ##

if ($office.text -replace "Users : .*" -eq "Actual Office Location ") {$path2 = "String that Dynamic DG needs"} else {$path2 = $office.text -replace "Users : .*"}

$path = $office.text -replace ".*: "
 
 
## Defines the veriable to use as text for the manager field ##

$managerdefined                  = $Manager.Text


## Check for if user already exists ##

$User      = "$($GivenName.text).$($Surname.text)"
$UserCheck = Get-ADUser -Filter {SamAccountName -eq $User}


If ($UserCheck -ne $Null) {


## Prompt -User already exists ##

Add-Type -AssemblyName System.Windows.Forms 
$UserError                       = New-Object system.Windows.Forms.Form
$UserError_L                     = New-Object System.Windows.Forms.Label
$UserError.Text                  = "Oops!"
$UserError.ClientSize            = '1120,189'
$UserError.Controls.Add($UserError_L)
$UserError.Visible               = $True
$UserError.Icon                  = "\\SMB path to icon"


$UserError_L.width               = 50
$UserError_L.height              = 10
$UserError_L.location            = New-Object System.Drawing.Point(59,79)
$UserError_L.Text                = "The UPN/Sam Account Name you are trying to use in already in use - Please try a different combination"
$UserError_L.AutoSize            = $True
$UserError_L.Font                = 'Microsoft Sans Serif,16'


Sleep -seconds 5


$UserError.Close()

}

## If the user doesn't exist - Create User ##

Else {

New-ADUser -Credential $UserCredential -Name "$($GivenName.text) $($Surname.text)" `
-DisplayName "$($GivenName.text) $($Surname.text)"`
-GivenName "$($GivenName.text)"`
-Surname "$($Surname.text)"`
-UserPrincipalName "$($GivenName.text).$($Surname.text)@Domain.co.nz"`
-SamAccountName "$($GivenName.text).$($Surname.text)"`
-EmailAddress "$($GivenName.text).$($Surname.text)@Domain.co.nz"`
-path "$path" `
-AccountPassword (ConvertTo-SecureString -AsPlainText "$Password" -Force) `
-Enabled $true `
-Description $Description.text `
-MobilePhone $Mobile.text -OfficePhone $TelephoneNumber.text `
-Office $path2 `
-OtherAttributes @{pager=$Mobile.text} `
-HomePage "www.Github.com" `
-Title $Description.text `
-Department $Department.Text `
-StreetAddress $StreetAddress.Text `
-City $City.Text `
-State $State.Text `
-PostalCode $PostCode.Text `
-Company "Company Name" `
-Manager "$managerdefined" `
-Country "NZ"


## Creates a remote mailbox both on prem and O365 ##

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://FQDN-OnPrem-Exchange.co.nz/PowerShell/ -Authentication Kerberos -Credential $UserCredential -ErrorAction SilentlyContinue
Import-PSSession $Session -DisableNameChecking -ErrorAction SilentlyContinue -AllowClobber

Enable-RemoteMailbox "$($GivenName.text) $($Surname.text)" -RemoteRoutingAddress "$($GivenName.text).$($Surname.text)@Domain.mail.onmicrosoft.com"


## User successfully created form ##

Add-Type -AssemblyName System.Windows.Forms 
$UserCreated                     = New-Object system.Windows.Forms.Form
$UserCreated_L                   = New-Object System.Windows.Forms.Label
$UserCreated.Text                = "Success!"
$UserCreated.ClientSize          = '925,189'
$UserCreated.Controls.Add($UserCreated_L)
$UserCreated.Visible             = $True
$UserCreated.Icon                = "\\SMB to icon"


$UserCreated_L.width             = 50
$UserCreated_L.height            = 10
$UserCreated_L.location          = New-Object System.Drawing.Point(59,79)
$UserCreated_L.Text              = "User Account + Mailbox successfully created - Don't forget to license the user in AD"
$UserCreated_L.AutoSize          = $True
$UserCreated_L.Font              = 'Microsoft Sans Serif,16'


Sleep -seconds 4


$UserCreated.Close()

}

$MainForm.Close()
})



## Admin Logs ##

$AdminLogs                       = New-object System.Windows.Forms.TextBox
$AdminLogs.Multiline             = $True;
$AdminLogs.location              = New-Object System.Drawing.Point(35,350)
$AdminLogs.Size                  = New-Object System.Drawing.Size(1370,230)
$AdminLogs.Font                  = 'Microsoft Sans Serif,10'
$AdminLogs.Scrollbars            = 3#Scrollbars.Vertical



 ## Admin Logs Label - (Admin Logs) ##

$AdminLogs_L                     = New-Object system.Windows.Forms.Label
$AdminLogs_L.text                = "Admin Logs (Only applicable when creating multiple users)"
$AdminLogs_L.AutoSize            = $true
$AdminLogs_L.width               = 25
$AdminLogs_L.height              = 10
$AdminLogs_L.location            = New-Object System.Drawing.Point(35,330)
$AdminLogs_L.Font                = 'Microsoft Sans Serif,10'




## Create user button + Label in one ##

$Button2                         = New-Object -TypeName System.Windows.Forms.Button
$Button2.text                    = "Create User + add another"
$Button2.width                   = 220
$Button2.height                  = 55
$Button2.location                = New-Object System.Drawing.Point(1190,275)
$Button2.Font                    = 'Microsoft Sans Serif,12'
$Button2.add_Click({

## Takes Office OU + Display name - Strips out Display name and leading space ##

if ($office.text -replace "Users : .*" -eq "Actual Office Location ") {$path2 = "String that Dynamic DG needs "} else {$path2 = $office.text -replace "Users : .*"}

$path = $office.text -replace ".*: "
 

## Defines the veriable to use as text for the manager field ##

$managerdefined        = $Manager.Text


## Check for if user already exists ##

$User      = "$($GivenName.text).$($Surname.text)"
$UserCheck = Get-ADUser -Filter {SamAccountName -eq $User}


If ($UserCheck -ne $Null) {


## Prompts that user already exists ##

Add-Type -AssemblyName System.Windows.Forms 
$UserError                       = New-Object system.Windows.Forms.Form
$UserError_L                     = New-Object System.Windows.Forms.Label
$UserError.Text                  = "Oops!"
$UserError.ClientSize            = '1120,189'
$UserError.Controls.Add($UserError_L)
$UserError.Visible               = $True
$UserError.Icon                  = "\\SMB share to icon"


$UserError_L.width               = 50
$UserError_L.height              = 10
$UserError_L.location            = New-Object System.Drawing.Point(59,79)
$UserError_L.Text                = "The UPN/Sam Account Name you are trying to use in already in use - Please try a different combination"
$UserError_L.AutoSize            = $True
$UserError_L.Font                = 'Microsoft Sans Serif,16'


Sleep -seconds 5


$UserError.Close()

}

## If the user doesn't exist - Create User ##

Else {

New-ADUser -Credential $UserCredential -Name "$($GivenName.text) $($Surname.text)" `
-DisplayName "$($GivenName.text) $($Surname.text)"`
-GivenName "$($GivenName.text)"`
-Surname "$($Surname.text)"`
-UserPrincipalName "$($GivenName.text).$($Surname.text)@Domain.co.nz"`
-SamAccountName "$($GivenName.text).$($Surname.text)"`
-EmailAddress "$($GivenName.text).$($Surname.text)@Domain.co.nz"`
-path "$path" `
-AccountPassword (ConvertTo-SecureString -AsPlainText "$Password" -Force) `
-Enabled $true `
-Description $Description.text `
-MobilePhone $Mobile.text -OfficePhone $TelephoneNumber.text `
-Office $path2 `
-OtherAttributes @{pager=$Mobile.text} `
-HomePage "www.Github.com" `
-Title $Description.text `
-Department $Department.Text `
-StreetAddress $StreetAddress.Text `
-City $City.Text `
-State $State.Text `
-PostalCode $PostCode.Text `
-Company "Company name" `
-Manager "$managerdefined" `
-Country "NZ"


## User information displayed in Admin logs ##

$Userspecified                   = "$($GivenName.text).$($Surname.text)"
$UserInfo                        = Get-ADUser -Filter {SamAccountName -eq $User} -Properties `
                                   DisplayName,GivenName,Surname,`
                                   UserPrincipalName,SamAccountName,EmailAddress,`
                                   Description,`
                                   MobilePhone,Office,`
                                   Title,Department,`
                                   StreetAddress,City,State,PostalCode,Manager



$users                           = $UserInfo.DisplayName,
                                   $UserInfo.EmailAddress,
                                   $UserInfo.Description,
                                   $UserInfo.MobilePhone,
                                   $UserInfo.Office,  
                                   $UserInfo.Department,
                                   $UserInfo.StreetAddress,
                                   $UserInfo.City,
                                   $UserInfo.State,
                                   $UserInfo.PostalCode,`
                                   $UserInfo.Manager


foreach($user in $users){
    $AdminLogs.AppendText($user)
    $AdminLogs.AppendText("`r`n")
    }
$AdminLogs.AppendText("`r`n")
$AdminLogs.AppendText("============================================================================================================")
$AdminLogs.AppendText("`r`n")




## Creates a remote mailbox both on prem and O365 ##

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://FQDN-OnPrem-Exchange.co.nz/PowerShell/ -Authentication Kerberos -Credential $UserCredential -ErrorAction SilentlyContinue
Import-PSSession $Session -DisableNameChecking -ErrorAction SilentlyContinue -AllowClobber

Enable-RemoteMailbox "$($GivenName.text) $($Surname.text)" -RemoteRoutingAddress "$($GivenName.text).$($Surname.text)@Domain.mail.onmicrosoft.com"


## User successfully created form ##

Add-Type -AssemblyName System.Windows.Forms 
$UserCreated                     = New-Object system.Windows.Forms.Form
$UserCreated_L                   = New-Object System.Windows.Forms.Label
$UserCreated.Text                = "Success!"
$UserCreated.ClientSize          = '925,189'
$UserCreated.Controls.Add($UserCreated_L)
$UserCreated.Visible             = $True
$UserCreated.Icon                = "\\SMB share to icon"


$UserCreated_L.width             = 50
$UserCreated_L.height            = 10
$UserCreated_L.location          = New-Object System.Drawing.Point(59,79)
$UserCreated_L.Text              = "User Account + Mailbox successfully created - Don't forget to license the user in AD"
$UserCreated_L.AutoSize          = $True
$UserCreated_L.Font              = 'Microsoft Sans Serif,16'


Sleep -seconds 4


$UserCreated.Close()




## If User creation was successful - Form is cleared for another user ##

$GivenName.Text                  = ''
$Surname.text                    = ''
$Description.text                = ''
$TelephoneNumber.text            = ''
$Mobile.text                     = ''
$Department.text                 = 'Start typing Department..'



$password                        = Get-RandomCharacters -length 5 -characters 'abcdefghiklmnoprstuvwxyz'
$password                       += Get-RandomCharacters -length 2 -characters 'ABCDEFGHKLMNOPRSTUVWXYZ'
$password                       += Get-RandomCharacters -length 1 -characters '1234567890'
$password                       += Get-RandomCharacters -length 2 -characters '!"$%&/()=?}][{@#*+'
$password                        = Scramble-String $password 


$PasswordSecureString.text       = $Password
$Manager.text                    = 'Start typing name..'
$StreetAddress.text              = ''
$City.text                       = ''
$State.text                      = ''
$PostCode.text                   = ''
$office.text                     = 'Start typing office location..'

}


})


###### FORM Controls ---- Don't forget to add any new Forms ---- ######

$MainForm.controls.AddRange(@($General,$TelAcc,$Org,$Address,`
$GivenName,$GivenName_L,`
$Surname,$Surname_L,`
$Description,$Description_L,`
$TelephoneNumber,$TelephoneNumber_L,`
$Mobile,$Mobile_L,`
$Department,$Department_L,`
$Manager,$Manager_L,`
$StreetAddress,$StreetAddress_L,`
$City,$City_L,`
$State,$State_L,`
$PostCode,$PostCode_L,`
$Office,$Office_L,`
$MailNote_L, `
$Button,$Button2, `
$PasswordSecureString,$Password_L, `
$AdminLogs,$AdminLogs_L
))


## Visually allows the Form to be presented ##

[void]$MainForm.ShowDialog()
