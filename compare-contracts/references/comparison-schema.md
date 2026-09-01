# Comparison Schema

Use this schema as the semantic contract for analysis. Adapt fields to the contract type, but keep source evidence separate from normalized values and analyst judgments.

## Portfolio-Level Fields

Record:

- `analysis_title`
- `review_perspective`
- `as_of_date`
- `comparison_groups`
- `contract_ids`
- `currency_assumptions`
- `materiality_thresholds`
- `known_limitations`

## Contract Record

For each contract, capture:

```json
{
  "contract_id": "C001",
  "group": "Supplier A",
  "file": "agreement.pdf",
  "document_type": "master services agreement",
  "parties": [],
  "effective_date": null,
  "initial_term": null,
  "renewal": null,
  "governing_law": null,
  "status": "active|expired|draft|unknown",
  "extraction_quality": "high|medium|low",
  "notes": []
}
```

## Normalized Term Record

Represent each comparison point as an object instead of a bare string:

```json
{
  "category": "payment",
  "term": "invoice_due",
  "raw_value": "within thirty (30) days",
  "normalized_value": 30,
  "unit": "calendar_days",
  "basis": "invoice_date",
  "status": "stated|not_stated|not_found|unclear|not_applicable",
  "conditions": [],
  "exceptions": [],
  "source_ids": ["S-C001-004"],
  "verification": "original_checked"
}
```

## Recommended Categories

- Identity: parties, affiliates, guarantors, signatories, contract hierarchy.
- Scope: services, deliverables, territory, volume, exclusivity, dependencies.
- Commercial: fixed fees, variable fees, indexation, taxes, expenses, rebates, minimum commitments.
- Payment: invoice trigger, due date, withholding, disputes, late fees, set-off.
- Performance: milestones, acceptance, KPIs, SLAs, service credits, remedies.
- Duration: effective date, initial term, renewal, notice windows, holdover.
- Termination: convenience, cause, cure, insolvency, transition, exit charges.
- Risk allocation: liability cap, cap basis, exclusions, indemnities, insurance, force majeure.
- Information: confidentiality, data roles, security, breach notice, retention, audit.
- Intellectual property: background IP, foreground IP, license, ownership, publicity.
- Governance: reporting, change control, subcontracting, assignment, audit, records.
- Legal: warranties, compliance, governing law, venue, dispute escalation, priority of documents.

## Normalization Rules

1. Preserve the original expression in `raw_value`.
2. Convert only when the conversion is unambiguous; record the chosen unit and basis.
3. Keep one-time, recurring, usage-based, contingent, and pass-through charges separate.
4. Keep calendar days and business days separate.
5. Treat tax inclusion, currency, escalation, and minimum-volume assumptions as material.
6. Separate a general rule from its exceptions and conditions.
7. Do not infer that silence grants a right or removes a duty.
8. When several documents form one agreement, record document precedence and cross-references.

## Differences and Recommendations

For every material difference, record:

- the normalized fact for each contract or group;
- the factual delta;
- the affected party and operational consequence;
- severity with rationale;
- confidence and source IDs;
- a recommendation phrased as an action, not a legal conclusion.
