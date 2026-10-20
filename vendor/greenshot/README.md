# Bundled Greenshot

`portable.zip` is the unmodified official Windows portable pre-release identified in `dependency.json`.
The SHA-256 matches the digest published by the upstream GitHub release API. It is copied into build/publish output;
RelaxCat verifies it and expands it into its own local application-data cache when selecting Greenshot.
No machine-wide Greenshot installation is required. This full portable build is x64 and requires .NET Framework 4.8.

Greenshot is a separate executable. Its license is reproduced in `LICENSE.txt` and retained inside the original archive,
along with all upstream files/notices. Corresponding unmodified source is available at the exact source/sourceArchive
URLs in `dependency.json`. Redistribution must retain these notices and provide the corresponding source as required
by the included upstream license. No changes are made to Greenshot source or binaries.

To update, query the official `greenshot/greenshot` GitHub releases, choose the latest suitable Windows portable
pre-release (not an installer or MCP package), replace the archive, verify its published digest, update this manifest
and `GreenshotBundle` constants, then run the real provider smoke test.
