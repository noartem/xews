# @noartem/xews

`xews` is a small CLI for transferring files through Exchange Web Services by using attachments on the first draft message.

It is useful when you already have EWS access and want a simple way to upload, list, download, and clear files through Exchange without opening a mail client.

It can also synchronize one local folder through a dedicated draft-based channel. Sync sends changes as compressed batch attachments and removes each package after it is applied.

## Install

Install from a local checkout:

```bash
npm link
```

After that, the CLI is available globally as:

```bash
xews
```

Install from npm after publishing:

```bash
npm install -g xews
```

## Configuration

You can let `xews` create and open the config for you:

```bash
xews init
```

Or create `~/.config/xews/auth.json` manually:

```json
{
  "email": "user@example.com",
  "password": "your-password",
  "url": "https://exchange.example.com/EWS/Exchange.asmx"
}
```

Required fields:

- `email`
- `password`
- `url`

`xews init` opens the file in your default terminal editor. It first checks `VISUAL`, then `EDITOR`, and falls back to `vi`.

## Usage

Typical file transfer flow:

```bash
xews upload --file ./archive.zip
xews ls
xews download
xews sync --dir ./data --channel team-a
```

Show help:

```bash
xews --help
```

Upload one or more files to the first draft:

```bash
xews upload --file ./file1.txt --file ./file2.pdf
```

Download all attachments from the first draft into the current directory:

```bash
xews download
```

Download attachments and delete the draft after download:

```bash
xews download --delete
```

List attachments from the first draft:

```bash
xews list
xews ls
```

Clear all attachments from the first draft:

```bash
xews clear
```

Synchronize one local folder through a dedicated draft subject:

```bash
xews sync --dir ./tmp/a --channel local-test
```

Run a single sync cycle and exit:

```bash
xews sync --dir ./tmp/a --channel local-test --once
```

## Commands

- `init`: create `~/.config/xews/auth.json` if needed and open it in your default editor
- `upload`: attach one or more files to the first draft
- `download`: download all file attachments from the first draft into the current directory
- `list` / `ls`: show attachment names and sizes from the first draft
- `clear`: remove all attachments from the first draft
- `sync`: synchronize one folder through a dedicated draft channel like `xews-sync:<channel>`

## Publish

Preview the package contents:

```bash
npm pack --dry-run
```

Publish to npm:

```bash
npm publish
```

For a scoped public package:

```bash
npm publish --access public
```
