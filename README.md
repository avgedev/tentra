# Tentra Outlook Add-in

Add-in Outlook qui ajoute un bouton « **Envoyer vers Tentra** » dans le ribbon des emails. Au clic, l'email courant (sujet, expéditeur, destinataires, body HTML, pièces jointes) est posté à l'instance Tentra qui tourne en local sur `localhost:5180`. Tentra ouvre alors sa modale d'import — l'utilisateur sélectionne les passages à transformer en sous-tâches et valide.

## Pré-requis

- Tentra qui tourne (= `npm run dev` dans le repo principal, port 5180)
- Outlook Classic, Outlook on the Web ou Outlook desktop nouveau
- Compte admin Microsoft 365 (= pour sideloader le manifest custom)

## Architecture

```
Outlook (ribbon button)
    │
    ▼
Add-in taskpane (HTML/JS, hébergé HTTPS)
    │   POST /api/import-email-staged
    ▼
Vite dev server (localhost:5180)
    │   stocke staged-emails/<id>.json
    │   broadcast SSE
    ▼
Tentra (browser tab)
    │   reçoit event SSE
    │   ouvre la modale d'import
    ▼
Utilisateur valide → tâche créée
```

## Hosting (= étape obligatoire)

L'add-in nécessite des fichiers hébergés en HTTPS (= contrainte Office Add-ins). Plusieurs options :

### Option 1 — GitHub Pages (recommandé, gratuit)

Pré-configuré pour le repo `avgedev/tentra` (= URL `https://avgedev.github.io/tentra/`). Si tu utilises un autre user/repo, remplace manuellement dans `manifest.xml` (toutes les URLs) et regénère le `<Id>` (GUID).

1. Push le contenu de `outlook-addin/` à la racine du repo `tentra` (cf. section « Push initial » ci-dessous)
2. Sur GitHub : `Settings` → `Pages` → Source : `main` branch / root → Save
3. Attends ~30 s, vérifie l'URL `https://avgedev.github.io/tentra/manifest.xml` dans ton browser. Tu dois voir le XML.

### Option 2 — autre HTTPS (Vercel, Netlify, S3+CloudFront…)

Idem, juste pointer le manifest sur l'URL HTTPS finale.

### Push initial — sans clone (= upload depuis GitHub web)

Le plus simple si tu n'es pas à l'aise avec git en ligne de commande :

1. Sur ta page de repo GitHub `avgedev/tentra`, clique **`Add file`** → **`Upload files`**
2. Glisse-dépose tous les fichiers de `outlook-addin/` (y compris les sous-dossiers `src/` et `assets/`) depuis l'explorateur Windows
3. Vérifie que la structure dans GitHub matche celle de ton dossier local
4. Commit message : `Initial add-in upload` (ou ce que tu veux)
5. **Commit changes**

### Push initial — via git en CLI

Si tu préfères :

```powershell
cd "<repo>"
git remote add origin https://github.com/avgedev/tentra.git
git add .
git commit -m "Initial add-in upload"
git push -u origin main
```

## Sideload du manifest

### Outlook Classic (desktop Win32) — via le centre admin

1. Centre d'administration Microsoft 365 → Paramètres → Applications intégrées → Téléverser une application personnalisée
2. Choisir « Charger un fichier de manifeste »
3. Sélectionner le fichier `manifest.xml` modifié
4. Confirmer le déploiement (= tu choisis qui est éligible : toi seul / un groupe / l'orga)
5. Attendre 6-24 h pour la propagation (Microsoft fait du caching aggressive). Pour un déploiement immédiat sur ton compte personnel, redémarrer Outlook après l'upload.

### Outlook on the Web — sideload direct sur ton compte (test rapide)

1. Outlook web → ⚙ Paramètres → Voir tous les paramètres → Général → Compléments
2. « Mes compléments » → « Ajouter à partir d'un fichier »
3. Sélectionne `manifest.xml`
4. L'add-in apparaît dans le ribbon de tes mails (parfois après refresh F5)

Les sideloads via Outlook web propagent généralement aussi à Outlook Classic du même compte sous quelques minutes.

## Utilisation

1. Ouvre un email dans Outlook
2. Dans le ribbon, clique sur **« Envoyer vers Tentra »** (icône Tentra)
3. Le taskpane s'ouvre à droite, affiche un résumé du mail
4. Clique sur **« Envoyer vers Tentra »** dans le taskpane
5. Tentra (qui doit être ouvert dans un onglet) affiche immédiatement la modale d'import
6. Tu sélectionnes les passages à transformer en sous-tâches → Importer

Si Tentra n'est pas ouvert au moment de l'envoi, le staging est posé en attente côté serveur Vite. Au prochain démarrage de Tentra, la modale s'ouvrira automatiquement avec ce staging.

## Sécurité

V1 : pas de token / auth. Le serveur Vite (`localhost:5180`) n'écoute que sur `localhost`, donc seul l'utilisateur sur cette machine peut POSTer. Acceptable tant que :
- on n'expose pas le port sur le réseau
- on ne déploie pas sur plusieurs machines

À durcir avec un token partagé dans un header `X-Tentra-Token` si on étend le déploiement. Cf. note dans `memory/MEMORY.md` du projet principal.

## Troubleshooting

### « Mixed content blocked » dans la console du taskpane

Le browser bloque les fetch HTTPS → HTTP localhost. Edge WebView2 (= utilisé par Outlook Classic récent) considère localhost comme « potentially trustworthy » et autorise par défaut, mais des versions plus anciennes peuvent bloquer.

Solution : activer HTTPS sur Vite via `@vitejs/plugin-basic-ssl` (= self-signed cert), puis pointer le add-in sur `https://localhost:5180/...`. À documenter dans Tentra principal le jour où c'est nécessaire.

### Le bouton n'apparaît pas dans le ribbon

- Redémarre Outlook après le sideload
- Vérifie que le manifest.xml est valide (= utilise [le validateur Microsoft](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/troubleshoot-manifest))
- Vérifie que toutes les URLs `https://avgedev.github.io/tentra/...` ont été remplacées
- Vérifie que les fichiers PNG `assets/icon-*.png` sont accessibles HTTPS

### Vite reçoit le POST mais Tentra n'ouvre pas la modale

- Vérifie que Tentra est ouvert dans un onglet du browser
- Ouvre la console (F12) sur l'onglet Tentra → cherche des erreurs de l'EventSource
- Recharge l'onglet Tentra : la lecture initiale de `/api/staged-emails` rattrape les pendings

## Icônes

Les icônes `assets/icon-*.png` sont des copies du logo Tentra (`public/Tentra_logo.png` du repo principal). Pour V1 c'est une placeholder unique pour toutes les tailles — Outlook scale automatiquement, c'est imparfait mais suffisant. Si tu veux des icônes propres :

1. Génère via une lib de redimensionnement (Photoshop, online tools, ou PowerShell + ImageMagick)
2. Tailles attendues : 16×16, 32×32, 64×64, 80×80, 128×128 px
3. Format PNG, transparence supportée

## Évolutions possibles

- **Sécurité** : token partagé Tentra ↔ add-in
- **Multi-instance** : surcharger l'URL de l'endpoint Tentra via un paramètre URL (`?tentra=https://...`) pour pouvoir cibler une instance Tentra distante
- **1-clic ExecuteFunction** : bouton qui POST directement sans ouvrir le taskpane (= UX plus fluide)
- **Indicateur visuel post-import** : marquer l'email avec une catégorie Outlook « Tentra » pour visualiser ce qui a déjà été importé
