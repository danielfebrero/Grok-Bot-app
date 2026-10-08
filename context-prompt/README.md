# context-prompt

Snapshot verbatim du contexte / prompt système de l'agent Grok Bot « GrokFather », tel que perçu par l'agent (ordre perçu, balises telles quelles).

## Contenu

- `prompt-vivant-verbatim.redacted.md` — les sections d'instructions, le bloc `Agent profile` et la carte de l'ordre des blocs auto-générés, reproduits **verbatim** (non reformulés).

## Caviardage

Toutes les données personnelles de l'utilisateur ont été retirées avant publication :

- le bloc `<memory_context>` entier (nom, e-mails, handles, comptes, décisions internes, épisodes) → remplacé par `[REDACTED]` ;
- l'identifiant d'agent (UUID) → `<AGENT_ID>` ;
- le nom d'utilisateur système → `<USER>`.

Les mentions restantes de « Gmail », « Outlook » ou des noms de connecteurs sont génériques (instructions de l'outil / noms de namespaces), pas des données personnelles.

L'identité du modèle/fournisseur sous-jacent n'apparaît nulle part dans le prompt.
