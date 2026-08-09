# AO Paysage — projet de développement (phase 1)

Ce dossier est le point de départ du vrai projet, à reprendre dans **Claude Code**
pour continuer le développement. Il reprend la même stack que
[planning-menage](https://github.com/Vialles/planning-menage) : React + Firebase
(Firestore + Auth), déployable sur Render.com comme votre app de ménage.

`src/App.jsx` contient le prototype fonctionnel construit en discussion — équipe,
planning Gantt, mémoire technique, suivi financier, export Excel. Les données y
sont encore en mémoire locale (`useState`), pas persistées. C'est la première
tâche de dev ci-dessous.

## Mise en route

1. **Installer les dépendances**
   ```
   npm install
   ```

2. **Créer le projet Firebase**
   - Aller sur https://console.firebase.google.com → "Ajouter un projet"
   - Activer **Authentication** → méthode Email/Mot de passe
   - Activer **Firestore Database** → mode production, région `eur3` (Europe)
   - Dans Paramètres du projet → Vos applications → ajouter une app Web,
     copier les identifiants

3. **Configurer les variables d'environnement**
   ```
   cp .env.example .env.local
   ```
   Renseigner les valeurs Firebase copiées à l'étape précédente.

4. **Déployer les règles de sécurité Firestore**
   Le fichier `firestore.rules` isole déjà les données par organisation
   (base du multi-comptes prévu en phase 3). À déployer via la Console
   Firebase ou la CLI (`firebase deploy --only firestore:rules`).

5. **Lancer en local**
   ```
   npm run dev
   ```

6. **Déployer** — sur Render.com comme planning-menage, ou Firebase Hosting.

## Feuille de route phase 1 (MVP interne)

- [ ] Remplacer l'accès partagé par une vraie connexion (Firebase Auth)
- [ ] Persister les projets/équipe/planning/finances dans Firestore
      (remplacer les `useState` de `App.jsx` par des lectures/écritures Firestore)
- [ ] Brancher une première source de veille réelle : l'API ouverte du BOAMP
      (données disponibles sur https://boamp-datadila.opendatasoft.com),
      filtrée sur des mots-clés paysage/aménagement
- [ ] Garder les autres plateformes de la Sources en liens manuels pour l'instant
      (AWS-Achat, e-marchespublics, profils régionaux) — pas d'API publique
      unifiée, veille manuelle plus fiable qu'un scraping fragile
- [ ] Déployer une première version utilisable en interne

## Phases suivantes (rappel)

- **Phase 2** : synchronisation Google Agenda (OAuth), élargissement des
  sources de veille, taux horaire/marge par projet
- **Phase 3** : architecture multi-comptes pour la revente à d'autres
  paysagistes, facturation, onboarding — à ne lancer qu'après validation
  de la demande auprès de quelques confrères
