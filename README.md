#  pyDMARC-to-ELK

pyDMARC-to-ELK is a python tool that connects to a mailbox and reads the DMARC reports and send the data to elastic.

## Setup

This tool assumes that only dmarc reports are in the mailbox.
You need to first create the config file.
You can do this with:

python3 writedefaultconf.py

After that you can edit the Settings/config.ini file and enter the correct credentials.



### Azure AD 

The setup assumes that you already have a shared mailbox setup. (eg. mailauth-reports@example.com)

Create a new app registration in your tenant. (eg. dmarcToElk)
Single tentant an no redirect URI is needed.

Under "API permisssions" grant "Mail.Read" & "Mail.ReadWrite" in the "Microsoft Graph" api. (Application permissions)

Grant the admin consent.


Under "Certificates & secrets" creat a new client secret an copy the value. (this is the 'secret' in the mail config).

Create a security group (eg. SG-dmarcToElk) and add the shared mailbox user to this group.


Follow the commands below to restrict the access to only the shared mailbox
```
 $restrictedGroup = New-DistributionGroup -Name "SG-dmarcToElk" -Type "Security" -Members mailauth-reports@example.com -Description "resitricted access for APP dmarc-to-elk"
 New-ApplicationAccessPolicy -AppId XXXXXXXXXXXX  -PolicyScopeGroupId $restrictedGroup.PrimarySmtpAddress -AccessRight RestrictAccess
 ```


Once the config file is set you can run the tool with:

python3 pyStart.py
