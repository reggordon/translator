# Translator Tool Privacy & Security Notice

## Purpose
This document outlines privacy and security considerations for using the Translator Tool in a company setting.

## Data Handling
- All terms submitted for translation are sent to external translation engines (GoogleTranslator, LibreTranslator) via the internet.
- Do NOT use this tool for sensitive, confidential, or proprietary information (e.g., customer data, internal codes, unreleased product names).
- Only technical terms and public-facing terminology should be processed.

## API & External Services
- Translation engines may log requests; review their privacy policies before use.
- No API keys are stored in the scripts. If using paid APIs, keep keys secure and never commit them to code repositories.

## File Storage
- Output files and logs are saved locally. Restrict access to these files and clean up logs if they contain sensitive information.
- Failed and suspect translation logs may include original terms; review before sharing or archiving.

## Dependency Management
- All dependencies are open source and widely used. Keep them updated to avoid vulnerabilities.

## Personal Data & Compliance
- The tool is not intended for processing personal data (PII) or customer information, minimizing GDPR/PII risks.
- If you must process sensitive data, consult your company’s privacy and compliance team first.

## Best Practices
- Use only for non-sensitive technical terms.
- Review privacy policies of translation engines.
- Restrict access to output files and logs.
- Never store or share API keys in code repositories.

## Contact
For privacy or security concerns, contact your company’s IT or compliance team.
