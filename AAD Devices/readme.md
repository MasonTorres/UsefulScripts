# Create-Devices

This sciprt will create a lot of device object in Azure AD. 1 every 2 seconds.

# Delete-Devices

This script does the following.
    - Get all devices from Azure AD
    - Creates a hashtable containing all **unique** device names. 
        - Each Unique device will contain one or more child obejct. These child objects represent each device with the same name.
    - Loops through each unique device.
        - Finds the child device with the most recent creation date and sets **DeviceStatus** to **Do Not Delete**
    - Optional code to export the hashtable of devices as a json file.
    - Optional code to export the hashtbale of devices as a csv file.
    - Loops through all unique devices and deletes all devices except for the one device marked **Do Not Delete**
        - Calls Microsoft Graph API in batches to speed up deletion of objects. 20 devices at a time will be deleted.