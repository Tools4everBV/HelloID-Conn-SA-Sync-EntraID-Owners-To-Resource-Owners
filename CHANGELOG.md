# Change Log

All notable changes to this project will be documented in this file. The format is based on [Keep a Changelog](https://keepachangelog.com/), and this project adheres to [Semantic Versioning](https://semver.org/).

## [2.0.0] - 2026-09-14

### Added

- Added support to run the script in VScode

### Changed

- **Breaking change:** Changed authentication from client secret to certificate based. Existing tasks must be updated with the new certificate global variables (`EntraIdCertificateBase64String` and `EntraIdCertificatePassword`).
- Renamed the project from HelloID-Conn-SA-Sync-Teams-Owners-To-Resource-Owners to HelloID-Conn-SA-Sync-EntraID-Owners-To-Resource-Owners
- Renamed the script from `Sync-Teams-Owners-To-Resource-Owners-Groupmemberships.ps1` to `Sync-EntraID-Owners-To-Resource-Owners-Groupmemberships.ps1`
- Documented the breaking change and migration steps in the README

## [1.0.0] - 2025-10-25

This is the first official release of HelloID-Conn-SA-Sync-Teams-Owners-To-Resource-Owners.
