# FlashFM — Voyance en direct : synchro inscriptions

Automatise le traitement des inscriptions à l'émission « La Voyance en direct »
(Flash FM, Limoges) :

1. Récupère les nouvelles inscriptions depuis l'API BOCIR (GraphQL).
2. Les ajoute dans le Google Sheet de suivi de la saison.
3. Prépare un brouillon de mail de confirmation pour chacune, **sans jamais
   l'envoyer automatiquement** — l'envoi nécessite toujours une validation
   manuelle (CLI ou interface web).

Fonctionne en cron sur un VPS / Raspberry Pi, sans navigateur ni pilotage
d'interface distante.

## Stack

Node.js + TypeScript, SQLite (`better-sqlite3`) pour l'état de synchro,
`googleapis` pour Sheets, `nodemailer` pour l'envoi SMTP, `express` pour
l'interface de validation.

## 1. Installation

```bash
npm install
npm run build
cp .env.example .env
```

Remplis ensuite `.env` (voir sections suivantes pour où trouver chaque valeur).
Node 18+ requis (utilise le `fetch` natif).

## 2. Créer le compte de service Google

1. Va sur [Google Cloud Console](https://console.cloud.google.com/), crée un
   projet (ou réutilise un projet existant).
2. Active l'API **Google Sheets API** (menu *APIs & Services > Library*).
3. *APIs & Services > Credentials > Create Credentials > Service Account*.
   Donne-lui un nom (ex. `flashfm-voyance-sync`), pas besoin de rôle projet
   particulier.
4. Une fois le compte créé, ouvre-le, onglet **Keys > Add Key > Create new key
   > JSON**. Un fichier JSON se télécharge : c'est le fichier à fournir via
   `GOOGLE_SERVICE_ACCOUNT_JSON`.
5. Place ce fichier hors du repo (ex. `./credentials/google-service-account.json`,
   déjà ignoré par `.gitignore`) et pointe `GOOGLE_SERVICE_ACCOUNT_JSON` dessus.

## 3. Partager le Google Sheet avec le compte de service

1. Ouvre le fichier JSON téléchargé, repère le champ `client_email`
   (ressemble à `flashfm-voyance-sync@<projet>.iam.gserviceaccount.com`).
2. Ouvre le Google Sheet de suivi des inscriptions dans un navigateur.
3. Bouton **Partager**, colle cet email, donne le rôle **Éditeur**, envoie
   sans notification (ce n'est pas une vraie boîte mail).
4. Récupère l'identifiant du Sheet dans son URL :
   `https://docs.google.com/spreadsheets/d/<GOOGLE_SHEET_ID>/edit` → mets cette
   valeur dans `GOOGLE_SHEET_ID`.
5. Renseigne `GOOGLE_SHEET_TAB` avec le nom exact de l'onglet de la saison en
   cours (ex. `2026-2027`) — à changer dans `.env` chaque nouvelle saison.

L'app n'a besoin d'aucun autre accès Google (pas de Drive, pas de Gmail).

## 4. Credentials BOCIR

- `BOCIR_API_KEY` et `BOCIR_BEARER_TOKEN` : à récupérer depuis le back-office
  BOCIR (section API / intégrations du compte Flash FM).
- `BOCIR_FORM_ID` : identifiant du formulaire « La Voyance en direct ». Change
  si le formulaire est un jour recréé — ne le code nulle part ailleurs que
  dans `.env`.

## 5. Credentials SMTP (OVH)

- `SMTP_USER` = `pascal@flashfm.fr`
- `SMTP_PASSWORD` = mot de passe de la messagerie OVH associée à ce compte
  (ou mot de passe d'application si l'authentification à 2 facteurs est
  activée sur l'espace client OVH — se génère dans l'espace client OVH,
  section Emails).
- `SMTP_HOST` / `SMTP_PORT` sont déjà pré-remplis (`ssl0.ovh.net` / `465`),
  à ajuster seulement si OVH change son infrastructure.

## 6. Lancer en local

```bash
# Synchro BOCIR -> Sheet (ne fait jamais d'envoi de mail)
node dist/cli.js sync

# Lister les inscriptions en attente de validation
node dist/cli.js list

# Envoyer le mail de confirmation d'une inscription (id BOCIR)
node dist/cli.js send <id>

# Envoyer tous les mails en attente d'un coup
node dist/cli.js send-all

# Interface web de validation (aperçu des mails, exclusion, envoi, historique)
node dist/cli.js serve
# -> http://localhost:3000 (port configurable via WEB_PORT)
```

## 7. Planifier l'exécution quotidienne

Une exécution quotidienne suffit (le formulaire ne reçoit pas des centaines
d'inscriptions par jour). Le script appelé ne fait toujours que `sync` : il ne
déclenche jamais d'envoi de mail. La validation et l'envoi restent une action
manuelle, via `node dist/cli.js send`/`send-all` ou l'interface web.

### Linux / macOS (cron)

Exemple, tous les jours à 8h :

```cron
0 8 * * * /chemin/vers/flashfm-voyance/scripts/cron-sync.sh
```

Si tu veux garder l'interface web disponible en permanence (par ex. sur un VPS),
lance `node dist/cli.js serve` comme service systemd/pm2 séparé du cron.

### Windows (Planificateur de tâches)

1. Vérifie que Node.js est installé et dans le PATH (ouvre PowerShell, tape
   `node -v` : une version doit s'afficher).
2. Ouvre le **Planificateur de tâches** (touche Windows, tape "Planificateur
   de tâches", ou `taskschd.msc`).
3. *Créer une tâche de base...* :
   - Nom : `Synchro Voyance BOCIR`.
   - Déclencheur : **Tous les jours**, choisis une heure (ex. 8h00).
   - Action : **Démarrer un programme**.
     - Programme/script : chemin complet vers `scripts\cron-sync.bat`
       (ex. `C:\FlashFM\voyance\scripts\cron-sync.bat`).
     - Laisse "Démarrer dans" vide : le script se positionne lui-même dans le
       bon dossier.
4. Une fois la tâche créée, clic droit dessus > **Propriétés** :
   - Coche **Exécuter même si l'utilisateur n'est pas connecté** si tu veux
     que ça tourne même PC verrouillé (il faudra saisir le mot de passe du
     compte Windows une fois).
   - Onglet **Conditions** : décoche "Ne démarrer la tâche que si l'ordinateur
     est branché sur secteur" si c'est un PC fixe sans batterie.
5. Dans les **Options d'alimentation** Windows (Panneau de
   configuration > Options d'alimentation), passe le mode veille sur
   **Jamais** pour cette machine — sinon la tâche planifiée ne se déclenchera
   pas si le PC est en veille à l'heure prévue.
6. Teste : clic droit sur la tâche > **Exécuter**, puis vérifie
   `data\cron.log` et le Sheet.

Pour garder l'interface web accessible en permanence sur ce PC, lance
`node dist\cli.js serve` dans une fenêtre PowerShell que tu laisses ouverte,
ou crée une seconde tâche planifiée au démarrage de la session qui lance cette
commande.

## 8. Données et logs

- État de synchro + historique des inscriptions : `data/voyance.sqlite`
  (chemin configurable via `DB_PATH`).
- Logs d'exécution : `data/app.log` (`LOG_PATH`), et sortie du cron dans
  `data/cron.log`.
- Rien de tout ça n'est committé (voir `.gitignore`).

## Sécurité

- Aucun secret n'est codé en dur : tout passe par `.env` (jamais committé) et
  le fichier JSON du compte de service Google (jamais committé non plus).
- Aucun mail n'est jamais envoyé automatiquement par le job planifié — c'est
  la règle non négociable de cette app.
- L'app ne fait que **lire** BOCIR, **ajouter des lignes** dans le Sheet (elle
  ne modifie ni ne supprime jamais une ligne existante), et **envoyer des
  mails uniquement sur validation manuelle explicite**.
