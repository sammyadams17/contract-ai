# Bringing the engine to a new client

The engine is generic. What makes it draft well for a client is that client's playbook, and the
playbook is written from their lawyers' feedback. Plan for three or four review rounds on each
contract type before the drafts are good.

## 1. Agree the scope (week 0)

- Pick **one contract type** with a standard template and a steady flow of term sheets.
- Agree who reviews the drafts on the client side: a named lawyer with time each round.
- Agree what "good" means: the lawyer's comments are about judgement calls, not about terms that
  were missed or wrongly written.
- Settle data handling: where documents and outputs live, and which model deployment may
  read them.

## 2. Collect (week 1)

- The client's **template** for the contract type, in Word, with its numbering and fields
  intact.
- **Five to ten past term sheets** for that type, with the signed contracts that came out of
  them. The signed contracts are the evidence for writing rules. Never use them as a base.
- Any **precedents** the lawyers draft from, and their house style notes and drafting notes.
- The client's own particulars (registered office, notice details) for `--client-details`.

## 3. Write the first playbook (week 1)

- Start from `playbooks/general.md` and copy `playbooks/example-nda/` as a skeleton.
- Add one `##` section for each clause and schedule of the template, named as the template names
  it.
- Read the past term sheets against the signed contracts. Every place the lawyers changed the
  template is a candidate rule. Write it in the section it touches.
- Write `vocabulary.md` from the client's term sheets: their names for things, and their
  abbreviations.
- See `playbook-guide.md` for how to write rules.

## 4. Draft and review (weeks 2–4, repeated)

1. Run `read` on a term sheet and check `record.json`: every term there, every quote real. Fix
   the vocabulary if the reader missed whole kinds of term.
2. Run `draft`, then `annotate`. Send the lawyer the redline and the annotated term sheet
   together. The annotated term sheet shows them where each term landed.
3. Take every comment the lawyer makes and turn it into a playbook rule (see the guide). Keep a
   log of comment → rule, so it's clear why each rule exists.
4. Redraft the affected pieces, then the whole contract, and check nothing else moved with
   `compare`.
5. Repeat with a new term sheet each round, so the rules generalise instead of fitting one deal.

## 5. Hand over

- The playbook is now the client's drafting knowledge in writing. Give its owner on the client
  side edit rights and the playbook guide.
- Set `DRAFTING_LLM_TOKENS_PER_MINUTE` and the worker counts for their model deployment, so
  drafts slow down rather than fail when the limit is reached.
- Agree how new templates and new contract types get added: each is a new playbook, following
  the same loop.

## Watch out for

- **Rules that fit one deal.** If a rule names one deal's figures or parties, it's wrong. Write
  the rule in general terms and let the term sheet supply the values.
- **Rules that fight.** A General rule and a section rule that disagree. The drafter accounts
  for both and satisfies neither. Search the playbook before adding a rule.
- **Template edge cases in Word.** Unusual numbering, text boxes or content controls. Fix them in
  `draft/writer.py` with a test that uses a small made-up document, never in the playbook.
- **Grading your own output.** The lawyer decides whether a draft is good. A score computed by
  the engine about its own draft isn't evidence.
