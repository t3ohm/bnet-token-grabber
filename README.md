# Battle.net Token Grabber

A tool to automate Battle.net login and extract authentication tokens.

![Application Screenshot](https://github.com/t3ohm/bnet-token-grabber/blob/master/LOOKATME.png)

## Features

- Automated Battle.net login process
- Token extraction from redirect URLs
- GUI interface with status monitoring
- Force logout functionality
- CLI support for headless operation

## Usage

### GUI Mode
1. Run the executable
2. Click "Start Login" to start the process
3. Copy the token to clipboard
### CLI Mode
```bash
TokenGrabber.exe -e "your@email.com" -p "password" --execute
```

Requirements
- Windows OS(not tested in linux)
- Internet Explorer (for automation)
- Battle.net account credentials
