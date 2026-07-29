# Switch Network Storage: Microsoft → Google

Scripts to move cloud storage from **Microsoft OneDrive** to **Google Drive**.

## On your Windows PC (run this)

Open **PowerShell as Administrator** and run:

```powershell
Set-ExecutionPolicy -Scope Process Bypass -Force
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/tymarler/txblt/cursor/network-storage-google-7115/scripts/switch-to-google-drive.ps1" -OutFile "$env:TEMP\switch-to-google-drive.ps1"
& "$env:TEMP\switch-to-google-drive.ps1"
```

Or clone this repo and run `scripts\switch-to-google-drive.ps1` locally.

### What the Windows script does

1. Disables OneDrive folder backup (Desktop, Documents, Pictures)
2. Resets Windows default save locations to local disk (`C:`)
3. Moves redirected folders back from OneDrive paths
4. Installs Google Drive for desktop (if missing)
5. Launches Google Drive sign-in

After it runs, sign in with your Google account and choose **Stream** or **Mirror** files.

## On Linux / Cursor Cloud Agent

```bash
./scripts/setup-google-drive-linux.sh
```

This installs `rclone`, configures Google Drive, and mounts it at `~/google-drive`. You will need to complete Google OAuth when prompted.

## Cursor MCP (Google Drive access in agent)

`.cursor/mcp.json` is configured for the Google Drive MCP server. Add your OAuth credentials at `.cursor/gdrive-credentials.json` after setting up the [Google Drive MCP server](https://github.com/modelcontextprotocol/servers/tree/main/src/gdrive).
