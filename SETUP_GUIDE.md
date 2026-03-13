# OpenClaw Port Setup Guide

If you ever need to set up OpenClaw from scratch or if the connection drops again, please refer to these settings.

## ⚠️ Core Concept: Two Separate Ports
OpenClaw uses **two** different ports that must not be mixed up:
1. **Gateway Port (`18792`)**: The main "brain" of the agent.
2. **Relay Port (`18795`)**: The "eyes" that look at your Chrome Browser.

*(Note: The Relay Port is automatically calculated by OpenClaw based on the Gateway Port. For a Gateway of 18792, the Relay becomes 18795).*

## 1. System Configuration File (`openclaw.json`)
Location: `C:\Users\307984\.openclaw\openclaw.json`
Make sure these two blocks exist:
```json
  "gateway": {
    "port": 18792,
    "mode": "local"
  },
  "browser": {
    "enabled": true
  }
```

## 2. Dashboard Settings (Web Page)
* **Relay Port**: `18795`
* **Gateway Access**: `ws://127.0.0.1:18792`

## 3. Chrome Extension Settings (Footprint Icon)
1. Right-click the footprint icon in the top right of Chrome -> **Options**
2. Set **Relay Port** to `18795`.
3. Click **Save & Test**.
4. Left-click the footprint icon -> **Attach current tab**.
