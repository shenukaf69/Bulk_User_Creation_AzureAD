# Bulk_User_Creation_AzureAD
Bulk_User_Creation_AzureAD

# bulk_User_Creation_Final.ps1

PowerShell 7 script to:

- Bulk create users in Microsoft Entra ID (Azure AD) from a CSV file  
- Assign Microsoft 365 **E1/E3** licenses plus **Teams** license  
- Optionally enable **archive mailbox** and **auto-expanding archive**  
- Export summary and skipped-user reports with a log file

> ⚠️ This script is intended for admins. Do **not** commit real passwords or production CSVs to GitHub.

---

## Prerequisites

- **PowerShell 7+**
- Modules:
  - [`Microsoft.Graph`](https://learn.microsoft.com/graph/powershell/installation) (v2+)
  - [`ExchangeOnlineManagement`](https://learn.microsoft.com/powershell/exchange/connect-to-exchange-online-powershell)

Install:

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
Install-Module ExchangeOnlineManagement -Scope CurrentUser

