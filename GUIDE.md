# Omniroute pour Word

Ce projet ajoute un volet Word qui envoie une question a Omniroute, affiche la reponse et peut la remplacer a l'emplacement du curseur. Le bouton **Utiliser la selection** recupere le texte actuellement selectionne dans Word et le place dans la question.

## Contexte automatique du document

Le volet lit automatiquement le texte du document Word courant et la sélection active à chaque clic sur **Envoyer**. Tu peux donc demander directement « corrige les fautes », « analyse ce document » ou « améliore ce passage » sans recoller le contenu dans la zone de question.

Le texte est envoyé au proxy Render puis à Omniroute pour traiter la demande. Pour éviter de dépasser les limites des modèles gratuits, le contexte est limité aux 50 000 premiers caractères du document. Pour un document plus long, sélectionne la partie à traiter ou demande une analyse par section.

## 1. Securite de la cle API

La cle Omniroute n'est pas dans le code du navigateur. Le fichier `server.js` agit comme proxy : il recoit la demande du volet, ajoute la cle dans `Authorization`, puis appelle Omniroute. La cle doit etre configuree dans Render sous **Environment > Environment Variables** avec le nom `OMNIROUTE_API_KEY`.

La valeur par defaut de `OMNIROUTE_API_URL` est l'endpoint fourni : `http://13.62.104.102:20128/v1/chat/completions`. Il est possible de definir cette variable dans Render si l'endpoint change.

## 2. Deployer sur Render

1. Poussez ce projet dans un depot Git. Ne committez jamais la cle API.
2. Dans Render, creez un **Web Service** connecte au depot.
3. Utilisez `Node` comme environnement.
4. Commande de build : `npm install && npm run build`.
5. Commande de demarrage : `npm run start:render`.
6. Ajoutez `OMNIROUTE_API_KEY` dans les variables d'environnement Render. Ajoutez aussi `OMNIROUTE_API_URL` uniquement si vous voulez remplacer l'URL par defaut.
7. Le service sert le dossier `dist/`, qui contient le volet, les scripts, les images et `manifest.xml`.

Le dossier publie est donc `dist/`. L'application complete doit toutefois etre deployee comme Web Service, car `server.js` sert aussi le proxy `/api/chat`.

## 3. Mettre a jour le manifeste

Apres la creation du service Render, copiez son URL HTTPS, par exemple `https://mon-service.onrender.com`.

Dans `appPackage/manifest.xml`, remplacez toutes les occurrences de `https://REMPLACER-PAR-URL-RENDER.onrender.com` par cette URL, sans ajouter de chemin apres le nom d'hote. Relancez ensuite :

```text
npm run build
```

Le fichier pret a installer sera `dist/manifest.xml`. Le manifeste XML est celui a utiliser avec Word 2024 Click-to-Run sans abonnement Microsoft 365.

## 4. Tester localement

Pour compiler :

```text
npm install
npm run build:dev
```

Pour lancer le service de production localement, configurez une cle dans le terminal (ne l'enregistrez pas dans un fichier suivi par Git), puis lancez :

```text
npm run build
npm run start:render
```

Le proxy ecoute sur `http://localhost:3000` par defaut. Pour un vrai chargement dans Word, le manifeste doit utiliser une URL HTTPS accessible par Word ; l'URL HTTPS locale generee par le serveur de developpement Toolkit est uniquement destinee au debug.

## 5. Installer sur un autre PC Windows avec un catalogue approuve

1. Copiez `dist/manifest.xml` et le dossier `dist/assets` dans un dossier partage reseau accessible par les autres PC, par exemple `\\SERVEUR\OfficeAddins\Omniroute`.
2. Dans Word, ouvrez **Fichier > Options > Centre de gestion de la confidentialite > Parametres du Centre de gestion de la confidentialite > Catalogues de complements approuves**.
3. Ajoutez le chemin UNC du dossier partage, puis activez l'option qui autorise l'utilisation de ce catalogue.
4. Redemarrez Word.
5. Ouvrez **Insertion > Mes complements > Dossier partage** et ajoutez Omniroute.

Le XML doit conserver des URLs HTTPS publiques pour `SourceLocation`, les icones et les commandes. Le partage reseau sert a distribuer le manifeste ; il ne remplace pas l'hebergement Render du volet.

## Compatibilite Office 2024

Le manifeste JSON fourni par le modele Agents Toolkit vise le manifeste unifie en version preview et des builds Microsoft 365 recents. Pour Office 2024 Click-to-Run sans abonnement, utilisez le manifeste classique `manifest.xml` present dans ce projet. Il demande l'ensemble `WordApi 1.3` et la permission `ReadWriteDocument`.
