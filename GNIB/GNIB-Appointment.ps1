<#  
 
File Name: GNIB-Appointment.ps1
Version: 5.0
Author: Rafael Carneiro Machado
E-Mail: rafaelcarneiromachado@gmail.com
Web: https://www.linkedin.com/in/rafaelcarneiromachado/
 
.SYNOPSIS
  
        Made for people Working or Studying in Ireland who needs to schedule an appointment against Irish Burgh Quay Registration Office in order to request or renew the Irish Identification Document, GNIB (Garda National Immigration Bureau)
        Check slots availability for GNIB appointments based on a desired date, if one available is found, then the script fills the form up with the proper information, leaving to you the only task to submit it.


COMPANION LINKS AND INFORMATION
Links:
    https://burghquayregistrationoffice.inis.gov.ie/Website/AMSREG/AMSRegWeb.nsf/AppSelect?OpenForm
    https://harshp.com/dev/utils/gnib-appointments/

Options:
Categories:
    Other, Study, Work

SubCategories:
    Other:
        Lost Card, Stolen Card, Join family member, Permission Letter
    Study:
        PhD, Masters, Higher National Diploma, Degree, English Language Course, Second Level, Pre-Masters, Visiting Students, 3rd Level Graduate Scheme
    Work:
        Work Permit Holder, Hosting agreement, Working Holiday, Atypical Working Schemes, Invest or Start a Business, Visiting Academics, Doctor, 3rd Level Graduate Scheme

I have a GNIB card or I have been registered before:
    Yes:
        Renewal
    No:
        New
#>

#PARAMETERS, VARIABLES AND FIELS
$CAT = "Work" #Category
$SUBCAT = "Work Permit Holder" #Sub Category
$CONFIRMGNIB = "Renewal" #I have a GNIB card or I have been registered before
$GNIBNO = "123456" #GNIB Number
$GNIBEXPDATE = "01/02/2018" #Expiry date of current permission / GNIB card
$GIVENNAME = "Luke" #Given Name
$SURNAME = "Skywalker" #Surname
$BDAY = "01/01/1980" #Date of Birth: 
$NATIONALITY = "25" #Nationality (The number corresponds to the option position: 25 = Brazil)
$EMAIL = "luke@theforce.com" #Email
$FAMILYAP = "2" #Is this a family application (1 = Yes | 2 = No)
$PASSPORT = "1" #Do you have a Passport or Travel Document? (1 = Yes)
$PASSPORTNO = "FR123456" #Passport Number or Travel Document Number 
[datetime]$MaxDate = '03-28-2018 16:00:00.000' #Type in the same format as the example, a maximum date that fits your needs. The script will try to find the first available and closest slot over this date
$URL = "https://burghquayregistrationoffice.inis.gov.ie/Website/AMSREG/AMSRegWeb.nsf/AppSelect?OpenForm"
$URI = "https://burghquayregistrationoffice.inis.gov.ie/Website/AMSREG/AMSRegWeb.nsf/(getAppsNear)?openpage&cat=$CAT&sbcat=All&typ=$CONFIRMGNIB"

#CHECK FOR SLOTS AVAILABILITY BASED ON THE MAX DATE

$request = Invoke-RestMethod -Uri $URI
If ($request.slots[0].id -eq "") {
    Write-Host "No slots availables for the category ($CAT) and subcategory ($SUBCAT) selected"
}
Else {
        $slotID = $request.slots[0].id
        [datetime]$slotTIME = $request.slots[0].time -replace ' - ',' '
        If($slotTIME -lt $MaxDateDeserved) {
                Write-Output "Found the following slot: ID = $slotID | TIME = $slotTIME"

                $ie = New-Object -com InternetExplorer.Application
                $ie.visible=$True
                $ie.navigate($URL)
                while($ie.ReadyState -ne 4) {start-sleep -m 100}

                #Show Hidden DIVs
                $ie.document.getElementById("dvSubCat").Style.Display = "block"
                $ie.document.getElementById("dvRenew").Style.Display = "block"
                $ie.document.getElementById("dvDeclareRenew").Style.Display = "block"
                $ie.document.getElementById("dvDeclareCheck").Style.Display = "block"
                $ie.document.getElementById("dvPPNo").Style.Display = "block"

                #Fill Form Fields
                $ie.document.getElementById("Category").Value = $CAT
                while($ie.Busy) { Start-Sleep -Milliseconds 100 }
                $ie.navigate('javascript:resetSubCatFld();');
                $ie.document.getElementById("SubCategory").Value = $SUBCAT
                $ie.document.getElementById("ConfirmGNIB").Value = $CONFIRMGNIB
                $ie.document.getElementById("GNIBNo").Value = $GNIBNO
                $ie.document.getElementById("GNIBExDT").Value = $GNIBEXPDATE
                $ie.document.getElementById("UsrDeclaration").Checked = $True
                $ie.document.getElementById("GivenName").Value = $GIVENNAME
                $ie.document.getElementById("SurName").Value = $SURNAME
                $ie.document.getElementById("DOB").Value = $BDAY
                $ie.document.getElementById("Nationality").SelectedIndex = $NATIONALITY
                $ie.document.getElementById("Email").Value = $EMAIL
                $ie.document.getElementById("EmailConfirm").Value = $EMAIL
                $ie.document.getElementById("FamAppYN").SelectedIndex = $FAMILYAP
                $ie.document.getElementById("PPNoYN").SelectedIndex = $PASSPORT
                $ie.document.getElementById("PPNo").Value = $PASSPORTNO

                #Click Look for Appointment
                $ie.document.getElementById("btLook4App").setActive()
                $ie.document.getElementById("btLook4App").click()

                #Show Type of Date Selection DIV and select Closest Dates
                $ie.document.getElementById("dvSelectChoice").Style.Display = "block"
                $ie.document.getElementById("AppSelectChoice").Value = "S"
                $ie.document.getElementById("dvSelectSrch").Style.Display = "block"
                $ie.document.getElementById("btSrch4Apps").setActive()
                $ie.document.getElementById("btSrch4Apps").click()
                while($ie.Busy) { Start-Sleep -Milliseconds 100 }

                #Show Available Slots
                $ie.document.getElementById("dvAppOptions").Style.Display = "block"
                while($ie.Busy) { Start-Sleep -Milliseconds 100 }

                #Get and Select First Available Slot ID
                $request2 = Invoke-RestMethod -Uri $URI
                $slotID2 = $request2.slots[0].id
                [datetime]$slotTIME2 = $request2.slots[0].time -replace ' - ',' '
                Write-Host "Confirming slot information: ID = $slotID2 | TIME = $slotTIME2"
                $strJavaScript = "javascript:bookit('" + $slotID2 + "');"
                $ie.navigate($strJavaScript);
                while($ie.Busy) { Start-Sleep -Milliseconds 100 }
                $ie.document.getElementById("dvAppOptions").Style.Display = "block"

                #Submit Form
                while($ie.Busy) { Start-Sleep -Milliseconds 100 }
                $ie.document.getElementById("AppID").Value = $slotID2
                #$ie.document.forms[0].submit() #Commented due to recaptcha process
        }
}