# PIWeb System Skills

This repository contains 23 deployable system-level File Skill packages. It includes the original 21-package snapshot exported from the local PIWeb runtime on 2026-08-04 plus independently maintained additions.

## Skills

- `any-to-markdown`
- `contract-auditor`
- `compare-contracts`
- `data-analysis`
- `group-legal-contract`
- `data-to-docx`
- `data-to-pdf`
- `data-to-xlsx`
- `docx`
- `edge-tts`
- `frontend-design`
- `html-to-pdf`
- `huaxin-pptx`
- `image-toolkit`
- `lucanet-excel-data-extraction`
- `office-convert`
- `pdf`
- `pdf-ocr`
- `pdf-to-image`
- `pdf-toolkit`
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
