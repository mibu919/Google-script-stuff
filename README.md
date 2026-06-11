# 📊 Google Apps Script Automation

This directory contains cloud-native Google Apps Script (`.js`) automation designed to bridge the gap between Google Workspace, external HR data, and Mobile Device Management (MDM) platforms. 

By leveraging the native Google Apps Script environment, these tools run completely serverless, eliminating the need for local infrastructure or scheduled tasks while maintaining direct access to enterprise Google APIs.

### 📜 Highlighted Scripts

#### 1. `Meraki-MDM-Integration.js`
**Purpose:** Syncs Google Workspace directory data with Cisco Meraki MDM.
- **Why it's advanced:** Many organizations struggle to keep their MDM asset tags in sync with their user directory. This script utilizes Google Apps Script's `UrlFetchApp` to securely communicate with the Meraki REST API. It dynamically maps Google Workspace user data (Emails, Names, OUs) directly to iPad and device properties in the Meraki Dashboard, ensuring that asset assignment is always 100% accurate without manual data entry.

#### 2. `Employee-Offboarding.js`
**Purpose:** Automates the complex Google Workspace offboarding lifecycle.
- **Why it's advanced:** Employee offboarding is often a manual, error-prone checklist. This script interacts with the Google Admin SDK and Gmail APIs to fully automate the process. Upon execution, it suspends the user, resets their password, transfers Google Drive ownership to their manager, hides them from the Global Address List (GAL), and sets up automatic email forwarding/auto-replies, ensuring immediate and secure offboarding compliance.
