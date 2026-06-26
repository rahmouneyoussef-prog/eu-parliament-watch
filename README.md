# EU Parliament Watch — PE full recent coverage

Bot de veille **recent-only** pour le Parlement européen.

Objectif : surveiller les publications récentes des principales sources officielles du Parlement européen et envoyer une alerte Telegram lorsque le texte HTML, PDF, XML ou JSON contient un mot-clé défini dans `keywords.txt`.

## Important

Aucun outil externe ne peut garantir absolument 100 % de *tout* le site du Parlement européen, car les documents sont publiés via plusieurs systèmes et certaines pages peuvent changer de structure. Cette version vise la couverture la plus large raisonnable des **publications récentes accessibles publiquement** dans les sources officielles ciblées : plénière, commissions, DOCEO, votes/RCV, registre public, Open Data, think tank et pages de presse.

## Sources surveillées

### Plénière
- Agendas
- Documents de plénière
- Questions parlementaires
- Votes et listes de votes
- Résultats de votes
- Roll-call votes / RCV, PDF et XML si les liens sont publiés
- Procès-verbaux
- Textes adoptés

### Commissions
Toutes les commissions connues sont incluses : AFET, DROI, SEDE, DEVE, INTA, BUDG, CONT, ECON, FISC, EMPL, ENVI, SANT, ITRE, IMCO, TRAN, REGI, AGRI, PECH, CULT, JURI, LIBE, AFCO, FEMM, PETI, EUDS, HOUS.

Pour chaque commission, le bot surveille :
- documents récents ;
- recherche documents ;
- rapports ;
- projets de rapport ;
- avis ;
- projets d'avis ;
- amendements ;
- amendements budgétaires ;
- documents de travail ;
- agendas ;
- procès-verbaux ;
- documents de réunion ;
- votes en commission.

### Autres sources PE
- DOCEO / recent documents
- Registre public
- Open Data API, lorsque les endpoints répondent
- RSS / stay informed
- Press room
- Think Tank / research publications

## Fichiers

- `watch.py` : script principal.
- `keywords.txt` : mots-clés à surveiller.
- `requirements.txt` : dépendances Python.
- `.github/workflows/watch.yml` : exécution automatique toutes les 30 minutes.
- `data/seen.json` : historique des URL déjà vues, créé automatiquement.
- `output/*.xlsx` : exports Excel générés automatiquement.

## Telegram

Deux secrets GitHub doivent exister :

- `TELEGRAM_BOT_TOKEN`
- `TELEGRAM_CHAT_ID`

Le bot envoie :
- un résumé Telegram ;
- un fichier Excel joint lorsqu'il trouve des résultats.

## Excel

L'Excel contient :
- onglet `Matches` : tous les résultats ;
- onglet `Votes_RCV_matches` : sous-ensemble des documents classés comme votes / RCV ;
- onglet `Readme` : explication des colonnes.

## Fréquence

Le workflow GitHub est configuré ainsi :

```yaml
cron: "*/30 * * * *"
```

GitHub accepte cette planification, mais l'heure exacte d'exécution peut parfois être retardée.

## Personnaliser les mots-clés

Modifier `keywords.txt`, puis commit.

Conseil pour un rapport précis : ajouter la référence du rapport et la référence de procédure, par exemple :

```text
A10-0177/2026
2023/0142(NLE)
Euro-Mediterranean Aviation Agreement
Morocco
Western Sahara
```
