# office-ltsc-install-script
Auto install script for Office 2024 LTSC

[![Not by AI](https://github.com/user-attachments/assets/8fae7c39-edc2-4c7e-bf18-b30e482de15f)](https://notbyai.fyi/)

# Configuration
1. Download both the install and uninstall scripts from the repo as well as the uninstall.xml file
   - You can optionally also download the Batch Files for easy manual execution
2. Download the Office Deployment Tool (ODT) from the [Microsoft Website](https://www.microsoft.com/en-us/download/details.aspx?id=49117)
3. Create your Deployment Configuration in the [Microsoft Office Admin Center](https://config.office.com/officeSettings/configurations) and download it.
   - (See the [Microsoft Documentation](https://learn.microsoft.com/en-us/office/ltsc/2024/deploy) for detailed instructions)
4. Upload all the files into a Shared Environment (SMB, CIFS, NFS)
5. Configure the script variables in both powershell files as needed
*These steps may change based on the way you want to deploy this. Please see the Wiki for more detailed and specific instructions*
