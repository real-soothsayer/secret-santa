# 🎅 Secret Santa

Application web pour organiser des tirages au sort de Noël avec persistance des données via Firebase.

## Fonctionnalités

- ✨ **Créer un événement** - L'organisateur crée un événement Secret Santa avec un code unique
- 🎟️ **Inviter des participants** - Partagez le lien ou le code pour que les participants rejoignent
- 👥 **Gestion des participants** - Voir en temps réel qui a rejoint l'événement
- 🎲 **Tirage au sort** - L'organisateur effectue le tirage quand tout le monde a rejoint
- 🎁 **Révéler les attributions** - Chaque participant peut révéler à qui il doit offrir un cadeau
- 📱 **Responsive** - Fonctionne sur mobile et desktop

## Configuration Firebase

### 1. Créer un projet Firebase

1. Allez sur [Firebase Console](https://console.firebase.google.com/)
2. Cliquez sur "Ajouter un projet"
3. Suivez les étapes de création

### 2. Activer Firestore

1. Dans votre projet Firebase, allez dans "Firestore Database"
2. Cliquez sur "Créer une base de données"
3. Choisissez le mode "Production" ou "Test"
4. Sélectionnez une région

### 3. Configurer les règles de sécurité Firestore

Dans Firestore > Règles, utilisez ces règles :

```javascript
rules_version = '2';
service cloud.firestore {
  match /databases/{database}/documents {
    match /events/{eventId} {
      allow read: if true;
      allow create: if true;
      allow update: if true;
    }
  }
}
```

> ⚠️ **Attention** : Ces règles sont permissives pour la simplicité. Pour une application en production :
> - Ajoutez Firebase Authentication
> - Restreignez les mises à jour aux seuls organisateurs
> - Limitez la création d'événements aux utilisateurs authentifiés
> - Ajoutez des règles de validation des données

### 4. Ajouter une application Web

1. Dans les paramètres du projet, cliquez sur "Ajouter une application" > Web
2. Donnez un nom à l'application
3. Copiez la configuration Firebase

### 5. Configurer l'application

Dans `index.html`, remplacez la configuration Firebase par la vôtre :

```javascript
const firebaseConfig = {
    apiKey: "VOTRE_API_KEY",
    authDomain: "votre-projet.firebaseapp.com",
    projectId: "votre-projet-id",
    storageBucket: "votre-projet.appspot.com",
    messagingSenderId: "123456789",
    appId: "1:123456789:web:abcdef123456"
};
```

## Déploiement sur GitHub Pages

1. Allez dans **Settings** > **Pages**
2. Sélectionnez la branche `main`
3. Cliquez sur **Save**

Votre application sera disponible à : `https://VOTRE-USERNAME.github.io/secret-santa/`

## Utilisation

### Pour l'organisateur

1. Cliquez sur "Créer un événement"
2. Entrez le nom de l'événement et votre nom
3. Partagez le lien ou le code avec les participants
4. Attendez que tout le monde rejoigne (minimum 3 participants)
5. Cliquez sur "Effectuer le tirage au sort"

### Pour les participants

1. Ouvrez le lien partagé ou entrez le code de l'événement
2. Entrez votre nom
3. Attendez que l'organisateur effectue le tirage
4. Révélez votre attribution !

## Technologies utilisées

- HTML5 / CSS3 / JavaScript (Vanilla)
- Firebase Firestore (base de données en temps réel)
- GitHub Pages (hébergement)

## Licence

MIT