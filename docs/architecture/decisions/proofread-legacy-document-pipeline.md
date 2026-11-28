# Markdown document proofreading before legacy Office output

Implemented contract. Automated regression coverage and native host acceptance are separate; the tests below do not establish packaged or native-host behavior.

Markdown→DOCX keeps the existing two-stage render/validate output. For Markdown→DOC/RTF/ODT/WPS, the Application first renders DOCX privately, proofreads privately and then converts the validated DOCX to the requested format. The public output remains the original Markdown-owned result.

Proofreading failure or cancellation never runs the final Office converter or publishes private input. If the user has requested retained intermediates, the verified DOCX is passed as an explicitly declared auxiliary to the shared Runtime finalizer rather than copied by ad hoc filesystem operations. Existing route schemas, manifest identity and cancellation owner remain authoritative.

The proofread options are available for these document targets in GUI and CLI; they remain inapplicable to Markdown→PDF, spreadsheet routes and named actions. Native backend format validation and Word/WPS/LibreOffice checks remain required.
