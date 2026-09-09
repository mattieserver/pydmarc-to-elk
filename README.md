#  pyDMARC-to-ELK

pyDMARC-to-ELK is a python tool that connects to a mailbox and reads the DMARC reports and sends the data to OpenSearch.

Requires Python 3.10 or later and an OpenSearch 3.x cluster (the `opensearch-py` 3.x client).
The config keys are still named `[elk]` for backwards compatibility with existing config files.

## Setup

This tool assumes that only dmarc reports are in the mailbox.

Install the dependencies and create the config file:

```
./install.sh
```

After that you can edit the Settings/config.ini file and enter the correct credentials.
Running `python3 writedefaultconf.py` again is safe: it only fills in options that are missing.

### Configuration

`[email]`

| Option | Description |
| --- | --- |
| `tenant_id` | Directory (tenant) ID of the app registration |
| `client_id` | Application (client) ID of the app registration |
| `secret` | Client secret value |
| `mailbox_id` | The shared mailbox to read (eg. mailauth-reports@example.com) |
| `delete_processed` | `yes` deletes a mail once every attachment has been ingested |

The whole `[email]` section is required only for `./run.sh`; see Usage below.

`[elk]`

| Option | Description |
| --- | --- |
| `host` / `port` | OpenSearch endpoint |
| `mode` | `read` logs the documents to debug.log, `write` indexes them |
| `auth` | `yes` uses https with `user` / `password` |
| `user` / `password` | OpenSearch credentials, used when `auth` is `yes` |
| `verify_certs` | `no` disables TLS verification, only for a self-signed cluster |

Reports are indexed into `dmarc-index-YYYY-MM` in a single bulk request per report, with a
deterministic document id (`<org_name>-<report_id>-<record index>`), so re-ingesting the same
report updates the existing documents instead of duplicating them.

### Report formats and fields we do not ingest

Two aggregate report formats are handled transparently:

| Format | Namespace | Spec |
| --- | --- | --- |
| Legacy | none | [RFC 7489](https://www.rfc-editor.org/rfc/rfc7489.html) Appendix C |
| Current | `urn:ietf:params:xml:ns:dmarc-2.0` | [RFC 9990](https://www.rfc-editor.org/rfc/rfc9990.html) |

RFC 9990 sets `elementFormDefault="qualified"`, so every element is namespace-qualified.
`strip_namespace()` in `pyDMARCELKv2.py` normalises the tree to bare tags before extraction,
so the same code path parses both. The archived copy in `data/processed/` keeps its namespace.

Format differences already handled: `policy_published/pct` (RFC 7489) is replaced by
`policy_published/testing` in RFC 9990, which also adds `policy_published/discovery_method`.
All three are ingested; the one that does not apply to a given report is an empty string.

The following optional elements are **defined by the specs but not currently indexed**.


| Element | Defined in | Notes |
| --- | --- |  --- |
| `record/identifiers/envelope_to` | RFC 7489 + RFC 9990  | The RCPT TO domain. By far the most common gap. |
| `record/auth_results/dkim/human_result` | RFC 7489 + RFC 9990  | Free-text detail on a DKIM failure. |
| `policy_published/fo` | RFC 7489 + RFC 9990 | Failure-reporting options requested by the policy. |
| `policy_published/np` | RFC 9990 only  | Policy for non-existent subdomains. Reporters emit it in RFC 7489 reports too, ahead of the spec. |
| `record/row/policy_evaluated/reason/type` and `/comment` | RFC 7489 + RFC 9990  | Why the receiver overrode the policy: `local_policy` (2,700), `forwarded` (57), `mailing_list` (17), `trusted_forwarder` (8). This is what explains a disposition that contradicts `p`. Repeatable, so it would index as an array. |
| `report_metadata/error` | RFC 7489 + RFC 9990 |  Repeatable. |
| `report_metadata/generator` | RFC 9990 only |  Software that produced the report. |
| `record/auth_results/spf/human_result` | RFC 9990 only |  RFC 7489 defines `human_result` under `dkim` only. |
| `<extension>` (file level and record level) | RFC 9990 only  | Carries a URI naming the extension. |

Two further schema notes worth knowing when reading the data:

- `auth_results/dkim` is `maxOccurs="unbounded"` in both specs, and `auth_results/spf` is
  unbounded in RFC 7489 but `maxOccurs="1"` in RFC 9990. All six `auth_results-*` fields are
  therefore indexed as arrays. See "Multiple authentication results" below.
- `policy_evaluated/reason` is `maxOccurs="unbounded"` in both.

### Multiple authentication results

A record can carry several DKIM results because a message may carry several DKIM signatures
(RFC 6376), and under RFC 7489 several SPF results because SPF is evaluated for both the
`helo` and `mfrom` identities. Records group messages that share identical authentication
results, so this is never an artifact of aggregation.

Alignment is not represented in `auth_results` at all. Per RFC 7489 section 3.1.1, a message
is a DMARC pass "if any DKIM signature is aligned and verifies". The receiver has already
applied that rule, and the verdict is the single `row/policy_evaluated/dkim` value.

**Use `row-policy_evaluated-dkim` for pass/fail rates. The `auth_results-*` arrays are raw
per-signature verification results and are diagnostics only** - a `fail` in there is routine,
and a signature that verifies but is not aligned contributes nothing to the DMARC outcome.

### Azure AD 

The setup assumes that you already have a shared mailbox setup. (eg. mailauth-reports@example.com)

Create a new app registration in your tenant. (eg. dmarcToElk)
Single tentant an no redirect URI is needed.

Under "API permisssions" grant "Mail.Read" & "Mail.ReadWrite" in the "Microsoft Graph" api. (Application permissions)

Grant the admin consent.


Under "Certificates & secrets" creat a new client secret an copy the value. (this is the 'secret' in the mail config).

Create a security group (eg. SG-dmarcToElk) and add the shared mailbox user to this group.


Follow the commands below to restrict the access to only the shared mailbox
```
 $restrictedGroup = New-DistributionGroup -Name "SG-dmarcToElk" -Type "Security" -Members mailauth-reports@example.com -Description "resitricted access for APP dmarc-to-elk"
 New-ApplicationAccessPolicy -AppId XXXXXXXXXXXX  -PolicyScopeGroupId $restrictedGroup.PrimarySmtpAddress -AccessRight RestrictAccess
 ```

## Usage

Once the config file is set you can run the tool with:

```
./run.sh
```

Every processed report is also written to `data/processed/`. To re-ingest that folder
(for example after fixing the index mapping) run:

```
./run.sh reload
```

Both are equivalent to activating `venv` and running `python3 pyStartv2.py` or
`python3 pyStartReloadDataV2.py` directly.

**The reload needs only the `[elk]` section.** Microsoft Graph is set up on demand, so
re-ingesting `data/processed/` works on a host that has no mailbox credentials at all -
the `[email]` section can hold placeholders or be missing entirely. Only `./run.sh`
validates `[email]` and acquires a token, and it does so before reading any mail.
