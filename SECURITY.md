# Security Policy

## Supported versions

| Version | Supported |
|---------|-----------|
| 0.1.x   | Yes       |

## Reporting a vulnerability

If you discover a security vulnerability in OpenSheet Core, please report it responsibly.

**Do not open a public issue.** Instead, email **hi@nader.info** with:

- A description of the vulnerability
- Steps to reproduce
- The potential impact
- A suggested fix (if you have one)

You should receive a response within 48 hours. We will work with you to understand and address the issue before any public disclosure.

## Scope

OpenSheet Core parses untrusted XLSX files, which are ZIP archives containing XML documents. Security-relevant areas include:

- **XML parsing** — XXE (XML External Entity) injection, billion laughs attacks
- **ZIP handling** — zip bombs, path traversal in entry names
- **Memory safety** — the Rust core provides memory safety guarantees, but unsafe code and FFI boundaries are areas of attention

## Current protections

- Built in Rust, which provides memory safety by default
- Uses `quick-xml` which does not process DTDs or external entities by default
- ZIP extraction uses the `zip` crate with standard safety checks
