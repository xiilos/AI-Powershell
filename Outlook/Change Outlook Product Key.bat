#Open the Command Prompt as administrator. Navigate to Office installation folder by running this command:
cd C:\Program Files\Microsoft Office\Office16

#If Office 2019 / 2016 32-bit is installed on 64-bit Windows, type this command instead:
cd C:\Program Files (x86)\Microsoft Office\Office16

#Now, you can change your Office product key by running the command below:
cscript ospp.vbs /inpkey:new_product_key


#Finally, type the following command to immediately activate your copy of Office 2019 / 2016 installation:
cscript ospp.vbs /act