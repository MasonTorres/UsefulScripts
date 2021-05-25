# Create-Devices

This sciprt will create a lot of device object in Azure AD. 1 every 2 seconds.

# Delete-Devices

This script does the following.
1. Get all devices from Azure AD
2. Creates a hashtable containing all **unique** device names. 
    - Each Unique device will contain one or more child obejct. These child objects represent each device with the same name.
3. Loops through each unique device.
    - Finds the child device with the most recent creation date and sets **DeviceStatus** to **Do Not Delete**
4. Optional code to export the hashtable of devices as a json file.
5. Optional code to export the hashtbale of devices as a csv file.
6. Loops through all unique devices and deletes all devices except for the one device marked **Do Not Delete**
    - Calls Microsoft Graph API in batches to speed up deletion of objects. 20 devices at a time will be deleted.