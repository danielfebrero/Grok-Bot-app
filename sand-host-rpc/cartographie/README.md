# cartographie

Cartographie **annotée** des services et méthodes RPC Sand/Cursor exposés par le host
(`api2.cursor.sh`, protocole Connect), avec le verdict d'accès observé pour chaque méthode.

## Fichiers
| Fichier | Contenu |
|---|---|
| `services-methodes.md` | Tableaux lisibles par service (méthode, kind, HTTP, verdict d'accès, note IDOR) + résumé et surface IDOR. |
| `services-methodes.json` | Version machine : `services.<Service> = [ {method, kind, http, verdict, idor_watch, …} ]`. |
| `audit-2026-10-24.md` | Rapport d'audit du balayage (copie de `rpc-audit-rapport.md`). |
| `data/all-rpc-methods.json` | Catalogue brut : 460 triplets `[service, méthode, kind]` (extrait des bundles host). |
| `data/rpc_catalog.json` | Schéma protobuf pour le décodage des corps (`messages`, `enums`, `methods`). Scalaires protobuf-es, imbriqué `ref#Type`. Régénéré par `extract-rpc-catalog.py` depuis `host-main.cjs`. |
| `data/rpc-sweep-log.tsv` | Balayage lecture : `service · method · kind · http · outcome`. |
| `gen-cartographie.py` | Générateur : fusionne `data/` → `services-methodes.{md,json}`. |

## Régénérer
```bash
python3 gen-cartographie.py
python3 extract-rpc-catalog.py /opt/sand/sand-host/host-main.cjs   # ou un .cjs.gz
```

## Méthodologie
- **Lecture seule.** Seules les méthodes à **préfixe de lecture** (Get/List/Read/Resolve/Watch/Check/…)
  ont été **appelées** avec un body `{}`. Les **mutations** (Create/Send/Set/Delete/Save/Start/Report/Submit…)
  sont **cataloguées mais jamais appelées**.
- **Jeton sand_mobile** (`type: sand_mobile`), chargé via `../token-helper`. Il est **scopé « compte »** :
  accepté par `GetMe` / `AvailableModels` / `ListGrokBotAgents` mais **rejeté** (unauthenticated) par ~138
  méthodes (admin team/sandbox, BackgroundComposer, Automations, AppPublish, la plupart d'AiService).
- **IDOR responsable.** Les fenêtres IDOR (auth OK mais id fourni par le client) sont **signalées**, et
  testées uniquement avec des ids **fournis par l'opérateur**, jamais devinés. Verdict : `GrokBotService`
  **non vulnérable** (scopé owner, lecture ET écriture) — voir `audit-2026-10-24.md`.
- `InferenceService/Stream` est **serveur-à-serveur** : le host remplace l'`authorization` du client par
  un Bearer de service via `getGrokBotAccessToken` → **inatteignable** avec un jeton client.

## Note sur le périmètre
Le catalogue (460 méthodes) est le **sous-ensemble audité**. Le moniteur `../rpc-monitor` suit, lui, la
**totalité** des déclarations RPC trouvées dans les bundles (`@generated from rpc …`, superset) pour
détecter toute dérive. Les deux chiffres diffèrent volontairement.
