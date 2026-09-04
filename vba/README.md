# Outlook VBA version

This folder contains the sanitised public version of the VBA automation that replaced the original Zapier workflow.

The production macro ran inside **classic Outlook for Windows** and converted structured class-confirmation emails into appointments in the default Outlook calendar.

## What it does

For matching confirmation emails, the VBA code:

1. checks for the exact subject `Schedule update: class added`;
2. extracts the Class ID and `dd.mm.yyyy hh:mm` date/time with regular expressions;
3. applies the historical +1 hour adjustment;
4. skips classes whose adjusted start time is already in the past;
5. checks the default calendar for the same Class ID;
6. creates a 25-minute appointment;
7. sets a 15-minute reminder;
8. moves the successfully processed email into `Inbox\Processed`.

It can run from three entry points:

- Outlook startup;
- `Application_NewMailEx` while Outlook is running;
- the public `Run_ProcessClassEmails_Manual` procedure, which can be attached to a ribbon or toolbar button.

## Public version versus the historical code

The public source preserves the real workflow but is not a verbatim copy of the private production module.

Changes made for this repository:

- the real teaching mailbox name was replaced with the configurable `SOURCE_MAILBOX_NAME` constant;
- unrelated Outlook macros were excluded;
- the Inbox sweep now iterates **backwards** before moving processed messages;
- `dd.mm.yyyy hh:mm` is parsed explicitly with `DateSerial` / `TimeSerial` rather than depending on Windows locale settings;
- constants and helper procedures make the workflow easier to read;
- lightweight failures are written to the VBA Immediate window with `Debug.Print`.

The backwards Inbox iteration addresses a limitation in the original implementation. The historical version moved messages out of the same `Inbox.Items` collection it was enumerating with `For Each`; in live use, that occasionally meant a second manual Sweep was needed to catch everything.

## Setup

1. Use **classic Outlook for Windows**. The new Outlook client does not run VBA macros.
2. Open the VBA editor with `Alt+F11`.
3. Open the `ThisOutlookSession` module.
4. Copy the contents of [`ThisOutlookSession.bas`](ThisOutlookSession.bas) into that module.
5. Change:

```vb
Private Const SOURCE_MAILBOX_NAME As String = "Teaching Account"
```

to the display name of the Outlook mailbox that receives the confirmation emails.

6. Create a `Processed` folder directly under that mailbox's Inbox.
7. Save the VBA project and restart Outlook if you want the startup and new-mail event handlers to run.
8. Optionally add `Run_ProcessClassEmails_Manual` to the Outlook ribbon or Quick Access Toolbar for a one-click Sweep.

## Expected input

A minimal synthetic email matching the parser is available at [`../examples/class-added.sample.txt`](../examples/class-added.sample.txt).

The two required body fields are:

```text
Class ID: SAMPLE001
Date/time: 15.01.2025 09:00
```

With the historical +1 hour rule, that example becomes a calendar appointment from **10:00 to 10:25**, with a reminder at **09:45**.

## Known limitations

The public version intentionally retains several design choices from the original implementation:

- the VBA version uses a fixed `+1 hour` adjustment rather than the timezone-aware London-to-Madrid conversion used by Zapier;
- duplicate detection scans appointment subjects for the Class ID rather than storing a dedicated external identifier;
- cancellations were not automated and remained a manual calendar task;
- the workflow depends on classic Outlook being installed and configured locally.

Those trade-offs are part of the case study: the VBA replacement removed the recurring SaaS cost and added useful safeguards, but the cloud-based Zapier version handled timezone conversion more cleanly.

## Privacy

The source in this repository contains no real mailbox address, employer name, production Class ID, authentication identifier or production email content.
