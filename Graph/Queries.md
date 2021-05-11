# For those times you need to query Graph API and can't remember the filter

## Email, Proxy Addresses, Other Mails
https://graph.microsoft.com/beta/users?$filter=proxyAddresses/any(x:x eq 'smtp:paige.turner@coprtech4.com')

https://graph.microsoft.com/beta/users?$filter=mail eq 'paige.turner@coprtech4.com'

https://graph.microsoft.com/beta/users?$filter=otherMails/any(x:x eq 'paige.turner@coprtech4.com')