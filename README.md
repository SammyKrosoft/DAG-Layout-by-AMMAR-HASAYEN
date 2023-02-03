# DAG layout script by Ammar Hasayen

This is a DAG layout script for Exchange 2013-2016-2019 by Ammar Hasayen adapted for my needs.

The author is Ammar Hasayen, MVP, Cloud and Cyber Security Expert : https://me.ahasayen.com/

The original script is on his blog at this address : https://blog.ahasayen.com/dag-database-activation/

Thanks Ammar for your great work on this that still help Microsoft Engineers to please Microsoft's customers in 2023 !


# Usage

- The result is an HTML file saved on the script's directory that you can send on the e-mail address of your choice, and will look like the below:

<img width="332" alt="image" src="https://user-images.githubusercontent.com/33433229/216522894-2297a8eb-ea6b-4314-84b8-239dc920f8f3.png">


- To update the e-mail address you want to send the report to, either use the -eMailSender, -eMailRecipient, and -eMailServer while launching the script, or hard code the default values of the parameters $eMailSender, $eMailRecipient and $eMailServer on the parameters definition on lines 53, 54, 55 of the script, and then just launch the script without specifying -eMailSender/Recipient/Server parameters:

 ![image](https://user-images.githubusercontent.com/33433229/216522215-94e4f68e-e409-4192-9621-bc8a9547ecc0.png)

- Example to launch the script with the default SMTP e-mail values:

```powershell
.\DAGLayout.ps1
```

- Example to launch the script with your own SMTP e-mail valueS:

```powershell
.\DAG_Layout.ps1 -eMailSender "samdrey@canadadrey.ca" -eMailRecipient "samdrey@canadadrey.ca" -eMailServer "E2016-01.canadadrey.ca"
```



# To download the script

Either copy/paste the text from the "raw" script data from the repo, or right-click Save As on the below link:

[DAG_Layout.ps1 (right-click Save Link As)](https://raw.githubusercontent.com/SammyKrosoft/DAG-Layout-by-AMMAR-HASAYEN/main/DAGLayout.ps1)
