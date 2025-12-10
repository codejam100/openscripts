#!/usr/bin/env python3
"""
imap_to_mbox.py

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
"""
from __future__ import annotations

import argparse
import base64
import codecs
import imaplib
import mailbox as mailbox_mod
import os
import re
import sys
from email import message_from_bytes
from getpass import getpass
from typing import List, Optional, Sequence, Tuple

DEFAULT_BATCH = 100


def decode_imap_utf7(s: bytes | str) -> str:
    if isinstance(s, bytes):
        raw = s
    else:
        raw = s.encode("utf-8", errors="surrogateescape")
    for codec_name in ("imap4-utf-7", "utf-7", "utf-8", "latin-1"):
        try:
            return codecs.decode(raw, codec_name)
        except (LookupError, UnicodeDecodeError):
            continue
    return raw.decode("latin-1", errors="replace")


def sanitize_filename(name: str) -> str:
    safe = re.sub(r"[\\/]+", "__", name)
    safe = re.sub(r"[\x00-\x1F\x7F]+", "_", safe)
    safe = safe.strip()
    if not safe:
        safe = "mailbox"
    return safe


def parse_list_response(data_line: bytes | str) -> str:
    """
    Extract mailbox name from a LIST/LSUB response.

    Handles responses like:
      (\\HasNoChildren) "." INBOX
      (\\HasNoChildren \\Marked \\Sent) "." "Sent"
      () "." Drafts

    Strategy:
    1. Try to match pattern: ... ) "delimiter" name
       - capture the rest after the quoted delimiter as the mailbox name.
    2. If not matched, fall back:
       - if there are quoted tokens, prefer the last quoted token unless it's the delimiter ('.')
       - otherwise use the last space-separated token.
    """
    if isinstance(data_line, bytes):
        s = data_line.decode("utf-8", errors="surrogateescape")
    else:
        s = str(data_line)

    # 1) Look for pattern: ... ) "delimiter" <name>
    m = re.search(r'\)\s+"[^"]*"\s+(?P<name>.+)$', s)
    if m:
        name = m.group("name").strip()
        # remove surrounding quotes if present
        if name.startswith('"') and name.endswith('"') and len(name) >= 2:
            name = name[1:-1]
        return name

    # 2) If pattern not found, fall back to quoted tokens:
    quoted = re.findall(r'"([^"]*)"', s)
    if quoted:
        # If more than one quoted group, the last quoted may be the name (unless it's the delimiter '.')
        last = quoted[-1]
        if last != ".":
            return last
        # if last is the delimiter, and there's text after it, try to take the final token
    # 3) Fallback: last whitespace-separated token
    parts = s.split()
    if parts:
        candidate = parts[-1]
        # strip quotes
        if candidate.startswith('"') and candidate.endswith('"') and len(candidate) >= 2:
            candidate = candidate[1:-1]
        return candidate
    return s


def _filter_and_decode(raw_items: Sequence[bytes | str], debug: bool = False) -> List[str]:
    mailboxes: List[str] = []
    for item in raw_items:
        if item is None:
            continue
        try:
            raw_name = parse_list_response(item)
        except Exception:
            raw_name = item.decode("utf-8", "surrogateescape") if isinstance(item, bytes) else str(item)
        decoded = decode_imap_utf7(raw_name)
        decoded_str = str(decoded).strip()
        if decoded_str == "" or decoded_str.upper() == "NIL" or decoded_str == ".":
            if debug:
                print(f"  (filtered) LIST entry: {repr(raw_name)} -> {repr(decoded_str)}")
            continue
        mailboxes.append(decoded_str)
    seen = set()
    out: List[str] = []
    for m in mailboxes:
        if m not in seen:
            seen.add(m)
            out.append(m)
    return out


def try_list_variants(imap: imaplib.IMAP4, debug: bool = False) -> Tuple[List[str], List[Tuple[str, Optional[Tuple[str, Sequence[bytes | str]]]]]]:
    tried = []
    variants = [
        ("list_no_args", ("list", None, None)),
        ("list_doublequote_star", ("list", '""', '*')),
        ("list_empty_star", ("list", '', '*')),
        ("list_none_star", ("list", None, '*')),
        ("list_none_percent", ("list", None, '%')),
        ("lsub_doublequote_star", ("lsub", '""', '*')),
        ("lsub_empty_star", ("lsub", '', '*')),
        ("lsub_none_star", ("lsub", None, '*')),
    ]

    for name, (cmd, ref, pat) in variants:
        try:
            if cmd == "list":
                if ref is None and pat is None:
                    status, data = imap.list()
                elif ref is None:
                    status, data = imap.list(None, pat)
                else:
                    status, data = imap.list(ref, pat)
            else:
                if ref is None:
                    status, data = imap.lsub(None, pat)
                else:
                    status, data = imap.lsub(ref, pat)
        except Exception as e:
            status, data = "NO", [f"exception: {e}".encode("utf-8")]
        tried.append((name, (status, data)))
        if debug:
            print(f"Tried {name}: status={status}, raw data={data!r}")
        if status == "OK" and data:
            parsed = _filter_and_decode(data, debug=debug)
            if parsed:
                return parsed, tried
    return [], tried


def get_all_mailboxes(imap: imaplib.IMAP4, debug: bool = False) -> List[str]:
    mailboxes, tried = try_list_variants(imap, debug=debug)
    if mailboxes:
        return mailboxes
    if debug:
        print("No useful LIST/LSUB results were found. Raw attempts:")
        for name, result in tried:
            status, data = result if result is not None else ("", [])
            print(f"  {name}: {status} -> {data!r}")
    return []


def fetch_mailbox_to_mbox(imap: imaplib.IMAP4, mailbox_name: str, outdir: str, batch_size: int = DEFAULT_BATCH) -> None:
    print(f"Processing mailbox: {mailbox_name!r}")
    try:
        encoded_name = codecs.encode(mailbox_name, "imap4-utf-7")
    except Exception:
        try:
            encoded_name = codecs.encode(mailbox_name, "utf-7")
        except Exception:
            encoded_name = mailbox_name.encode("utf-8", errors="surrogateescape")

    try:
        select_name = encoded_name.decode("utf-8", "surrogateescape")
        typ, data = imap.select(f'"{select_name}"', readonly=True)
        if typ != "OK":
            typ, data = imap.select(select_name, readonly=True)
    except Exception:
        try:
            typ, data = imap.select(mailbox_name, readonly=True)
        except Exception as e:
            print(f"  WARNING: Could not select mailbox {mailbox_name!r}: {e}")
            return

    if typ != "OK":
        print(f"  WARNING: Could not select mailbox {mailbox_name!r}: {typ} {data}")
        return

    typ, data = imap.uid("search", None, "ALL")
    if typ != "OK":
        print(f"  WARNING: UID SEARCH failed for {mailbox_name!r}: {typ} {data}")
        return

    uid_list = []
    if data and data[0]:
        uid_list = data[0].split()
    total = len(uid_list)
    print(f"  {total} messages found")

    outpath = os.path.join(outdir, sanitize_filename(mailbox_name) + ".mbox")
    mbox = mailbox_mod.mbox(outpath)
    mbox.lock()
    try:
        if total == 0:
            mbox.flush()
            return

        for start in range(0, total, batch_size):
            batch_uids = uid_list[start:start + batch_size]
            uid_range = b",".join(batch_uids).decode("ascii", errors="ignore")
            typ, fetch_data = imap.uid("fetch", uid_range, "(RFC822)")
            if typ != "OK":
                print(f"  WARNING: fetch failed for UIDs {uid_range}: {typ}")
                continue
            for item in fetch_data:
                if not item:
                    continue
                if isinstance(item, tuple) and len(item) >= 2 and item[1]:
                    msg_bytes = item[1]
                    try:
                        msg = message_from_bytes(msg_bytes)
                        mbox.add(msg)
                    except Exception as e:
                        print(f"    ERROR parsing message bytes: {e}")
            mbox.flush()
            print(f"  fetched {min(start + batch_size, total)}/{total}")
    finally:
        mbox.unlock()
        mbox.close()


def main() -> None:
    parser = argparse.ArgumentParser(description="Download IMAP mailboxes to mbox (LOGIN + SSL/STARTTLS).")
    parser.add_argument("--server", required=True)
    parser.add_argument("--port", type=int, help="IMAP port (defaults: ssl=993, starttls=143)")
    parser.add_argument("--username", required=True)
    parser.add_argument("--password", help="IMAP password (will prompt if omitted)")
    parser.add_argument("--outdir", default="./mboxes")
    parser.add_argument("--batch", type=int, default=DEFAULT_BATCH)
    parser.add_argument("--security", choices=("ssl", "starttls"), default="ssl")
    parser.add_argument("--debug", action="store_true", help="Show server greeting, capabilities and raw LIST/LSUB attempts")
    parser.add_argument("--auth-plain", action="store_true", help="Try AUTHENTICATE PLAIN instead of LOGIN (diagnostic)")
    parser.add_argument("--mailboxes", help="Comma-separated mailbox names to download (bypass LIST). Example: INBOX,Sent")
    parser.add_argument("--timeout", type=int, default=60)
    args = parser.parse_args()

    password = args.password or getpass("IMAP password: ")

    os.makedirs(args.outdir, exist_ok=True)

    port = args.port if args.port else (993 if args.security == "ssl" else 143)
    print(f"Connecting to {args.server}:{port} (security={args.security}) ...")

    try:
        if args.security == "ssl":
            imap = imaplib.IMAP4_SSL(args.server, port)
        else:
            imap = imaplib.IMAP4(args.server, port)
            typ, data = imap.starttls()
            if typ != "OK":
                print(f"STARTTLS failed: {typ} {data}")
                try:
                    imap.logout()
                except Exception:
                    pass
                sys.exit(1)
    except Exception as e:
        print("Connection failed:", e)
        sys.exit(1)

    if args.debug:
        try:
            print("Server greeting:", getattr(imap, "welcome", None))
            print("Server CAPABILITY:", getattr(imap, "capabilities", None))
        except Exception:
            pass

    if args.auth_plain:
        try:
            def auth_callback(challenge: bytes) -> str:
                auth_str = "\0" + args.username + "\0" + password
                return base64.b64encode(auth_str.encode("utf-8")).decode("ascii")
            typ, data = imap.authenticate("PLAIN", auth_callback)
            if typ != "OK":
                print("AUTHENTICATE PLAIN failed:", typ, data)
                try:
                    imap.logout()
                except Exception:
                    pass
                sys.exit(1)
            print("AUTHENTICATE PLAIN succeeded")
        except imaplib.IMAP4.error as e:
            print("AUTHENTICATE PLAIN failed:", e)
            try:
                imap.logout()
            except Exception:
                pass
            sys.exit(1)
        except Exception as e:
            print("AUTHENTICATE attempt error:", e)
            try:
                imap.logout()
            except Exception:
                pass
            sys.exit(1)
    else:
        try:
            typ, msg = imap.login(args.username, password)
        except imaplib.IMAP4.error as e:
            if args.debug:
                print("Login failed:", e)
                print("Server CAPABILITY (post-login attempt):", getattr(imap, "capabilities", None))
            else:
                print("Login failed:", e)
            try:
                imap.logout()
        except Exception:
            pass
            sys.exit(1)
        print("Login OK")

    # If user provided mailbox names, use those and skip discovery
    if args.mailboxes:
        mailbox_list = [m.strip() for m in args.mailboxes.split(",") if m.strip()]
        if args.debug:
            print("Using user-specified mailboxes:", mailbox_list)
    else:
        mailbox_list = get_all_mailboxes(imap, debug=args.debug)

    if not mailbox_list:
        print("No mailboxes found. If you used --debug, check the raw LIST/LSUB output above.")
        try:
            imap.logout()
        except Exception:
            pass
        sys.exit(0)

    print(f"Found {len(mailbox_list)} mailboxes")
    for mbox_name in mailbox_list:
        try:
            fetch_mailbox_to_mbox(imap, mbox_name, args.outdir, batch_size=args.batch)
        except Exception as e:
            print(f"Error processing mailbox {mbox_name!r}: {e}")

    try:
        imap.logout()
    except Exception:
        pass

    print("Done. mbox files saved in:", os.path.abspath(args.outdir))


if __name__ == "__main__":
    main()