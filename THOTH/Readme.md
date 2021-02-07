
# THOTH - Ultimate System Inventory & Log Collector
----------------------

* With more and more end users working from home, remotely, the IT departments have started to face a new challenge when the same users are experiencing technical issues with their local computers;
* Sometimes a remote shared session is not possible due to the issue per-se, and end users are not capable to provide deep technical information in order to help on the troubleshooting process;
* This script has been built in a way that it can be directly shared with end-users, where the only thing they have to do, is to download and execute the **THOTH-FORM.PS1 as Administrator**. It will compile a simple GUI to instruct the end-user through the process; 
* The script will collect basic and advanced logs from the local computer, creates a HTML report and zips everything into a single file that can then, be shared with the user's IT department;
* Hopefully, the collected logs and data will give us enough information to, at least, re-establish remote connectivity where we can then, jump in to a remote shared session, in order to finish the break-fix process.

----------------------

> **If you just want to test/use the tool, you can just download the [THOTH.EXE](/THOTH.EXE) file**;
> **If you want to read, understand, modify, customize and improve the code, keep reading;**

----------------------

## Technical Information:
* **Results and files are placed at: C:\LogCollector**

### Customization & Branding
You can change the variables at the begining of the code in order to customize the look and feel of the HTML report.
Customizable items and their default values are:
* Title: THOTH
* Logo: An image of THOTH, God of Record Keeping (Embedded in the code via BASE64 image string)
* Footer: Created by Rafael Machado (rafaelcarneiromachado@gmail.com)
* Main Background Color = "#3B3B3B"
* Main Font Color = "#FFFFFF"

### Data Inventoried & Collected
1. Hardware Inventory;
2. Network Inventory, including ISP information and Internet Speed Test (Download Throughput and Average Ping);
3. Driver Inventory;
4. Windows Services Statuses;
5. Resultant Set of Policies (GPOs / GPResult);
6. Software Inventory;
7. Windows Update: Hotfixes Installed and Missing;
8. EventViewer Events from Application and System since last boot time;
9. Overall System Performance Report considering a sample of 60 seconds, taken during the data gathering process;
10. Mobile Device Management Logs (if the device is corporate managed via a MDM platform);
11. CBS Log;
12. SCCM Logs (If SCCM Agent installed);
13. VMWare Horizon Logs for Horizon Agent and Horizon Client (If installed);

### Notes:
**When reviewing the code, press CTRL + M to EXPAND/COLLPASE the PS Regions. That will make it easier to understand the code structure.**

* The [**THOTH-FORM.PS1**](bin/THOTH-FORM.PS1) has everything embedded, including the function [**INVOKE-THOTHSCAN.PS1**](bin/INVOKE-THOTHSCAN.PS1), which is the script that performs the actual scan;
* I've put that way in order to make it easier to send just one single file to end users or to compile the PS script into an EXE file;
* If you don't want to use the GUI, you can extract (copy/paste) the function code of INVOKE-THOTHSCAN and use it as a stand-alone PS function;

#### PS1 to EXE:
* I'd recommend to compile the Thoth-Form.ps1 into an .EXE file in order to make it more user-friendly;
* In order to do that, you can use Inno Setup: https://jrsoftware.org/isinfo.php
* The project page at GitHub has the InnoSetup script file (**THOTH.iss**), which you can import, modify as you wish, and compile/re-compile the PS1 into an EXE file


