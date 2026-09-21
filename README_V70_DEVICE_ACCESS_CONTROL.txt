# V70 — Device Access Control

Added browser/device registration for the UNITY PATH RECOVERY AND COLLECTION SERVICES OPC EMS.

## How it works

1. Every browser gets a unique Device ID stored in a secure HttpOnly cookie.
2. Employees can only log in from an `Active` device registration assigned to their employee account.
3. If an employee logs in from a new device, the EMS creates a `Pending` device and displays the Device ID.
4. Admin opens **Device Management**, selects the employee, and approves the device.
5. Blocked or unregistered devices cannot log in to employee accounts.
6. Admin devices can be bootstrapped/registered with valid admin credentials.

## Important limitation

This is a web-device/browser registration system, not a literal MAC-address filter. Normal browsers do not expose a laptop's MAC address to a web application.

The Device ID is tied to the browser's persistent cookie and user-agent. Clearing browser cookies, using a different browser/profile, or reinstalling the browser can create a new Device ID and require re-registration.

For stronger hardware-level control later, use a company VPN/device certificate/managed-device solution.

## Admin flow

- Log in as Admin.
- Open **Device Management**.
- Pending employee devices appear there.
- Assign the device to the correct employee.
- Click **Approve**.
- To immediately stop a computer, click **Block**.

No additional third-party subscription is required for this feature.
