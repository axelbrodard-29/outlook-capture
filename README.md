# Capture d'e-mail – complément Outlook

Ajoute un bouton **« Capturer l'e-mail »** dans le ruban d'un e-mail ouvert dans Outlook. La capture contient l'en-tête (objet, expéditeur, date, destinataires, pièces jointes) et le contenu. L'extension vous demande ensuite le format : **PNG, JPEG, PDF ou presse-papiers**.

## 1. Mettre en ligne sur GitHub Pages (environ 5 min)

1. Sur github.com, créez un dépôt **public** nommé `outlook-capture`.
2. Cliquez sur **Add file → Upload files** et déposez **tout le contenu de ce dossier** : `manifest.xml`, `taskpane.html`, `taskpane.js`, `taskpane.css`, `commands.html` et le dossier `assets`.
3. Allez dans **Settings → Pages**. Choisissez *Deploy from a branch*, puis la branche `main` et le dossier `/ (root)`, et enregistrez.
4. Une minute plus tard, le site est disponible à l'adresse `https://VOTRE-COMPTE.github.io/outlook-capture/`.

## 2. Mettre votre adresse dans le manifeste

Ouvrez `manifest.xml` dans un éditeur de texte (Bloc-notes suffit). Remplacez **toutes** les occurrences de `USERNAME` par votre nom de compte GitHub (Ctrl+H → Remplacer tout), puis enregistrez.
Ce fichier reste sur votre PC : c'est lui que vous installez dans Outlook. Vous n'avez pas besoin de le renvoyer sur GitHub.

## 3. Installer dans Outlook classique

1. Ouvrez la page https://aka.ms/olksideload dans votre navigateur (connectez-vous avec votre compte Sagerime).
   Autre chemin depuis Outlook : **Accueil → Obtenir des compléments** (ou **Toutes les applications → Ajouter des applications**).
2. Allez dans **Mes compléments → Compléments personnalisés → Ajouter un complément personnalisé → Ajouter à partir du fichier…**
3. Sélectionnez votre `manifest.xml` et confirmez.
4. Redémarrez Outlook, puis ouvrez un e-mail. Le bouton **Capturer l'e-mail** apparaît dans le ruban, sous l'onglet Accueil ou Message.

> Si l'option « Ajouter un complément personnalisé » n'apparaît pas, votre administrateur l'a désactivée. Demandez à l'IT de déployer `manifest.xml` depuis le **Centre d'administration Microsoft 365 → Paramètres → Applications intégrées → Charger des applications personnalisées**. Ils peuvent ainsi l'installer pour toute l'équipe.

## Utilisation

Ouvrez un e-mail → **Capturer l'e-mail** → un aperçu s'affiche → choisissez le format.

- **PNG / JPEG** : image haute résolution (deux fois la taille d'affichage).
- **PDF** : format A4, découpé sur plusieurs pages si l'e-mail est long.
- **Presse-papiers** : à coller directement dans Teams, Word, un autre e-mail… (Ctrl+V)

Les fichiers sont nommés automatiquement : `Objet de l'e-mail_AAAA-MM-JJ`.

## Bon à savoir

- **Images externes** : les images intégrées (logos, signatures) sont bien capturées. En revanche, les images hébergées sur certains sites web peuvent manquer, car ces sites bloquent leur copie (règle de sécurité des navigateurs, appelée « CORS »).
- **Si le téléchargement ne démarre pas** dans votre version d'Outlook, faites un clic droit sur l'aperçu puis « Enregistrer l'image sous… », ou passez par le presse-papiers.
- **Version d'Outlook** : il faut un Outlook Microsoft 365 récent (moteur Edge WebView2). Les très anciennes versions d'Outlook 2016/2019 ne sont pas prises en charge.
- **Confidentialité** : la capture est créée localement, sur votre PC. Aucun contenu d'e-mail n'est envoyé à GitHub ni ailleurs. GitHub sert seulement à fournir les fichiers de l'extension.

## Contenu

| Fichier | Rôle |
|---|---|
| `manifest.xml` | Déclare le bouton et les adresses (fichier à installer dans Outlook) |
| `taskpane.html/.css` | Panneau de l'extension |
| `taskpane.js` | Code du complément + bibliothèques intégrées (html2canvas, jsPDF, DOMPurify) |
| `commands.html` | Fichier technique exigé par Outlook |
| `assets/` | Icônes |
