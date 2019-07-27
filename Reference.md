# Word Rhyme

* [TypeScript](http://www.typescriptlang.org/).
* [Office add-in documentation](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
* [OfficeDev Add-in samples on Github](https://github.com/officedev)

## Debugging

This template supports debugging using any of the following techniques:

- [Use a browser's developer tools](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-in-office-online)
- [Attach a debugger from the task pane](https://docs.microsoft.com/office/dev/add-ins/testing/attach-debugger-from-task-pane)
- [Use F12 developer tools on Windows 10](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-using-f12-developer-tools-on-windows-10)

## Local Testing

### Install Packages

> npm install

### Build the project

> npm run build

### Set up the Server

1. Install localhost certificates
    > npm run localhost-certificates-install

    * Make sure to accept the prompts

1. Check certificates were installed

    > npm run localhost-certificates-verify

### Run the Server

> npm run server