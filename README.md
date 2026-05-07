# TU89 Collection — ระบบรับของตามเลขที่สมาชิก

A Google Apps Script web app for managing member item collection (bag, polo shirt, jersey) for the Triamudom Family TU89 members.

## Features

- Look up members by 4-digit membership number (`89-XXXX`)
- Display item details (size, gender, jersey text/number)
- Toggle collected status per item type (bag / polo / jersey)
- Deny list — members with `denied` in column AG are blocked from collection
- 60-second server-side cache to reduce sheet reads

Column headers are looked up by name for jersey/polo size fields.

## License

GNU Affero General Public License v3.0 — see [LICENSE.md](LICENSE.md)
