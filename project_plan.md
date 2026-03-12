# API Token Inventory Project – Planning Guide

This is a solid infrastructure governance initiative. Here's how I'd approach it systematically across your four clusters (prod, ccs, qa, dev).

---

## Phase 1: Discovery (Enumerate All Tokens)

The first step is getting a complete picture of what exists before you try to document anything.

**For each cluster, you'll query the Elasticsearch API:**

```bash
# List all API keys in a cluster
GET /_security/api_key?with_limited_by=true

# Or via curl
curl -s -u elastic:$PASS https://<cluster>:9200/_security/api_key?with_limited_by=true | jq .
```

This returns every API key with its ID, name, creator username, creation time, expiration, and invalidation status. You'll want to run this against all four clusters and export to JSON/CSV for analysis.

**Key fields to capture per token:**
- `id`, `name`, `username` (creator), `realm`
- `creation`, `expiration`, `invalidated`
- `role_descriptors` (the actual permissions)
- `limited_by` (parent user's roles)

---

## Phase 2: Enrichment (Fill in Context)

Raw API output won't tell you *who uses it or why* — that's the hard part. For each token, you'll need to cross-reference:

- **Logstash configs** — look for `api_key:` entries in pipeline configs across your nodes
- **Filebeat/Agent configs** — fleet policies, enrolled agents
- **Kibana saved objects** — index patterns, connectors, alerting rules that use API keys
- **Jenkins/CodeBuild pipelines** — your `elastic_agents_cicd` pipeline likely has some
- **Application configs** — any app team configs stored in repos or Parameter Store
- **Watcher configs** — some watchers use API keys for actions or webhooks

For tokens where the creator is identifiable (a real username), reach out to that team directly to confirm purpose and owner.

---

## Phase 3: Documentation Structure

I'd recommend a spreadsheet with one row per token across all clusters. Suggested columns:

| Column | Notes |
|---|---|
| Cluster | prod / ccs / qa / dev |
| Token Name | From ES |
| Token ID | From ES |
| Creator Username | From ES |
| Created Date | From ES |
| Expiration | Flag if expired/no expiry |
| Status | Active / Invalidated |
| Permissions | Summarized from role_descriptors |
| Purpose | What it does |
| Team/Owner | Who owns it |
| Used In (Process/Script) | Pipeline name, script path, app name |
| Source Location | Logstash node, GitHub repo, CodeBuild, etc. |
| Stored in PMP | Yes/No + PMP entry name |
| Notes | Anything unusual |

---

## Phase 4: PMP Storage

For each active token's *value* (the actual secret), store it in ManageEngine PMP with a consistent naming convention, e.g.:

```
ES-PROD-logstash-ingest-apikey
ES-CCS-kibana-reporting-apikey
ES-QA-filebeat-fleet-apikey
```

Since you've already worked with the PMP REST API, you can automate bulk creation of these entries using your existing Python patterns. Each PMP entry should include the cluster, purpose, and owner in the notes/description field.

---

## Phase 5: Cleanup Pass

Once everything is documented, flag tokens that are:
- **Expired** — can be formally invalidated
- **Orphaned** — creator no longer on the team, no known usage
- **Overprivileged** — `all` indices or `superuser` role with no clear justification
- **Duplicated** — multiple tokens doing the same thing for the same team

---

## Suggested Execution Order

1. **Start with dev/qa** — lower risk, good for validating your discovery script
2. **Move to ccs** — cross-cluster search keys are especially important to document since they touch prod data
3. **Prod last** — most critical, most tokens, needs careful validation

---

## Tooling to Build

I can help you write a Python script that:
- Hits the `/_security/api_key` endpoint on all four clusters
- Normalizes the output into a flat CSV/Excel
- Flags expired, no-expiry, and invalidated tokens
- Optionally cross-references Logstash config directories for API key usage

Want me to start there, or would a spreadsheet template for the documentation phase be more useful first?
