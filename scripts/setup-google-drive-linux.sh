#!/usr/bin/env bash
# Setup Google Drive network storage on Linux (Cursor Cloud Agent / Ubuntu)
set -euo pipefail

MOUNT_POINT="${HOME}/google-drive"
RCLONE_REMOTE="gdrive"

echo "==> Installing rclone (if needed)"
if ! command -v rclone &>/dev/null; then
  curl -fsSL https://rclone.org/install.sh | sudo bash
fi

echo "==> Creating mount point: ${MOUNT_POINT}"
mkdir -p "${MOUNT_POINT}"

echo "==> Configuring rclone remote '${RCLONE_REMOTE}'"
if ! rclone listremotes 2>/dev/null | grep -q "^${RCLONE_REMOTE}:$"; then
  echo ""
  echo "  A browser window will open for Google sign-in."
  echo "  If you're on a headless machine, copy the URL shown and paste the token back."
  echo ""
  rclone config create "${RCLONE_REMOTE}" drive scope drive
fi

echo "==> Testing Google Drive connection"
rclone about "${RCLONE_REMOTE}:" || {
  echo "ERROR: Could not connect to Google Drive."
  echo "Run: rclone config reconnect ${RCLONE_REMOTE}:"
  exit 1
}

echo "==> Mounting Google Drive at ${MOUNT_POINT}"
if mountpoint -q "${MOUNT_POINT}" 2>/dev/null; then
  echo "  Already mounted."
else
  # FUSE mount in background
  nohup rclone mount "${RCLONE_REMOTE}:" "${MOUNT_POINT}" \
    --vfs-cache-mode writes \
    --daemon \
    --log-file "${HOME}/.config/rclone/mount.log"
  sleep 2
fi

if mountpoint -q "${MOUNT_POINT}"; then
  echo ""
  echo "========================================="
  echo " Google Drive mounted at: ${MOUNT_POINT}"
  echo "========================================="
  ls -la "${MOUNT_POINT}" | head -10
else
  echo "Mount may still be starting. Check: tail -f ${HOME}/.config/rclone/mount.log"
fi
