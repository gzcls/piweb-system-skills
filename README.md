# PIWeb System Skills

This repository is a deployment snapshot of the 21 active system-level File Skills exported from the local PIWeb runtime on 2026-08-04.

## Skills

- `any-to-markdown`
- `contract-auditor`
- `data-analysis`
- `data-to-docx`
- `data-to-pdf`
- `data-to-xlsx`
- `docx`
- `edge-tts`
- `frontend-design`
- `html-to-pdf`
- `image-toolkit`
- `lucanet-excel-data-extraction`
- `office-convert`
- `pdf`
- `pdf-ocr`
- `pdf-to-image`
- `pdf-toolkit`
- `pptx`
- `qrcode`
- `video-toolkit`
- `xlsx`

## Production installation

Install through PIWeb's governed File Skill ingestion endpoint. Do not copy these directories directly into the production runtime projection.

```http
POST /api/skills/install
Content-Type: application/json

{
  "package": "https://github.com/gzcls/piweb-system-skills.git",
  "scope": "system"
}
```

The production installation creates new installation and revision identities. Compatibility findings must be reviewed in the File Skill approval queue before affected packages become active.

Runtime dependencies such as LibreOffice, Chromium, FFmpeg, Python packages, and global Node packages belong in the production sandbox image; they are not provisioned by this repository.

## Licensing

Some packages contain their own license files. Those licenses apply to their respective package contents. No repository-wide license is granted for files that do not include an explicit license.
