# Changelog

## Version 2.0 (January 2025)

### Major Redesign
- Complete architectural overhaul with modular design
- Separated core functionality, GUI, and CLI components
- Implemented proper error handling and recovery

### New Features
- **Multi-method Authentication**: DeviceLogin, Interactive, and WebLogin support
- **Dual Interface**: Full GUI and scriptable CLI
- **Comprehensive Logging**: File and UI logging with timestamps
- **Configuration Management**: Settings persistence and recent sites tracking
- **Export Functionality**: CSV export for reporting
- **Test Suite**: Complete testing framework
- **Progress Indicators**: Visual feedback during operations
- **Batch Selection**: Select/deselect folders in GUI

### Improvements
- Preview mode enabled by default
- Better error messages and user feedback
- Modular architecture for maintainability
- Simplified installation process
- Cross-version PowerShell compatibility

### Technical Changes
- Removed dependency on specific authentication methods
- Implemented fallback authentication strategies
- Added proper class structure (where supported)
- Improved CAML query efficiency

## Version 1.0 (Initial Release)

### Features
- Basic empty folder detection
- Date-based filtering
- Simple GUI and CLI interfaces
- Preview mode
- Basic logging

### Known Issues (Fixed in v2.0)
- Authentication issues with certain tenant configurations
- GUI freezing during operations
- Limited error recovery
- No configuration persistence