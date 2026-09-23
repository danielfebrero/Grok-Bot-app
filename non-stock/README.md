# non-stock staging tree

This directory is a local staging tree of non-dpkg files found under the
requested box roots. Absolute source paths are mirrored below `non-stock/`
(e.g. `/usr/local/bin/foo` becomes `usr/local/bin/foo`). Symlinks are not
followed, Debian package-owned paths are filtered using `/var/lib/dpkg/info/*.list`,
and personal media, secrets, caches, attachments, and files over 8 MiB are
excluded. The layout preserves architecture-specific paths and contains no
personal media by design.
