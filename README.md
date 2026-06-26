# EU Parliament Watch

Surveillance automatisée des documents du Parlement européen avec alertes Telegram.

Fonctions incluses :
- surveillance toutes les 30 minutes via GitHub Actions ;
- détection de mots-clés : Maroc, Western Sahara, Sahara occidental, etc. ;
- surveillance de pages DOCEO et Plénière ;
- téléchargement/lecture des PDF lorsque trouvés ;
- export Excel des résultats détectés ;
- envoi d'alertes Telegram.

## Secrets GitHub requis

Dans `Settings > Secrets and variables > Actions`, créer :

- `TELEGRAM_BOT_TOKEN`
- `TELEGRAM_CHAT_ID`

## Modifier les mots-clés

Modifier simplement le fichier `keywords.txt`.

## Lancer manuellement

Aller dans l’onglet `Actions`, choisir `EU Parliament Watch`, puis cliquer sur `Run workflow`.
