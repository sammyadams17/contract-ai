# Writing a playbook

The playbook is how a client's drafting knowledge reaches the agents. It is plain Markdown, kept
in `playbooks/`, and a lawyer can read and edit it without touching code. Better drafts mostly
come from a better playbook, not better code.

## Shape

```
@include ../general.md          the rules every contract follows

# <Client> <contract type> playbook

## General (<contract type>)    rules for every piece of this contract type
- ...

## <Section name>               rules for one clause, schedule or block of the template
- ...
Read also by: <Other section>; <Another>
```

- **`@include`** brings in another file from the same folder or a relative path, in its place.
  Every client playbook includes `general.md`.
- **`## General`** (and `## General (<type>)`, which is read as the same section) holds rules
  that reach every piece. Each bullet goes to every drafter, so keep it true wherever it lands.
- **Every other `##` section** is named after a part of the client's template: "Definitions",
  "Payment", "Term and Termination". The planner gives each piece of the template one section
  by name, and that section's text goes to the piece's drafter whole.
- **`Read also by:`** sends a section's text to the pieces of other sections too. Use it when
  one clause's rule changes another's words.
- A bullet is `- ` at the start of a line. Continuation lines are indented two spaces.
- An optional `vocabulary.md` beside the playbook tells the term sheet reader the client's
  synonyms and abbreviations (`drafting-engine read --playbook ...` picks it up).

## What goes where

| The rule is about | Put it in |
|---|---|
| Any contract, any client (defined terms, open points, brackets, figures) | `general.md` |
| Every piece of this client's contract type | `## General (<type>)` |
| One clause or schedule | That clause's section |
| How a term sheet word maps to the template's defined term | `## General (<type>)`, as an equivalence |
| What the client's words mean when reading a term sheet | `vocabulary.md` |

Change `general.md` only for a rule you would give every client. Anything that comes from one
client's house style goes in that client's playbook.

## Writing a rule that works

- **Say what the draft must do, in the template's terms.** "The survival period is the term
  sheet's; where it gives none, the template's stands." Not "handle survival correctly".
- **One rule, one point.** A bullet the drafter can tick off as met, not met, or not
  applicable. Every drafter has to account for every bullet it gets.
- **Name the trigger.** "Where the term sheet is one-way ..." A rule with no trigger is applied
  everywhere.
- **Say what happens when the term sheet is silent.** Usually the template stands, or a `[•]`
  with one open point.
- **Examples are illustrations.** A bracketed example from a past deal is read as an
  illustration, never as this deal's fact, but it can still lead the agent. Prefer made-up or
  generic examples.
- **No hidden logic in code.** If a fix needs a phrase list or a special case in code, it
  belongs in the playbook instead.

## Turning a lawyer's comment into a rule

1. Find the comment's cause: a rule that's missing, a rule that's wrong, or a rule the drafter
   didn't follow.
2. **Missing:** write the rule in the section of the piece it's about, or in General if it
   reaches every piece.
3. **Wrong:** rewrite it. Check every other section for a rule that says the opposite.
4. **Not followed:** make the rule's trigger and its "done" state plainer, rather than repeating
   it.
5. Redraft only the pieces it affects (`draft RUN --only "<title>"`) and check the comment is
   answered. Then redraft the whole contract once, so you can see nothing else moved
   (`drafting-engine compare old.docx new.docx`).
