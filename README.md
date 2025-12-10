Download IMAP mailboxes to mbox files.

This update improves LIST/LSUB parsing so mailbox names are extracted
correctly when servers return lines like:
  (\\HasNoChildren) "." INBOX
  (\\HasNoChildren \\Marked \\Sent) "." Sent

Features:
 - LOGIN with normal password
 - SSL or STARTTLS
 - --debug prints server greeting, CAPABILITY and raw LIST/LSUB attempts
 - --mailboxes to supply a comma-separated list of folders to download (bypass LIST)
 - skips obviously-invalid LIST results ('.', 'NIL', empty)

Usage:

python3 imap_to_mbox.py --server imap.yourmailserver.com --username you@yourmailserver.com --password a728x@b2l_ --debug --outdir ./ you@yourmailserver.com