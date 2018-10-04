# Outlook Contacts Sync

Syncs domain users from Active Directory to your Microsoft Outlook Contacts list

## Working

This tool will query your on-premise Active Directory for users and syncs the result with Microsoft Outlook. If your organization’s Active Directory is very large (1500+ users), you can filter the results in the settings screen.

## Example Active Directory Query

For example, if you like to filter on physical office location and departments, use a query like this
```
(|(physicaldeliveryofficename=SCHIPHOL*)(department=NL*)(department=EQ-NL*)(department=EQ-EQ-NL*)(company=NETHERLANDS))
```

## Main Screen

![alt Outlook Contacts Sync](https://raw.github.com/NavaTron/outlook-contacts-sync/master/Source/Windows%20Store/Images/Screen1.png)
