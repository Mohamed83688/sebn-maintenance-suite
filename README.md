# SEBN-TN Enterprise Maintenance Management System

> **Système Intégré de Maintenance Industrielle et Automobile — SEBN-TN**  
> *Développé par Mohamed Aissaoui*

---

## 📌 Aperçu du Projet

La solution logicielle **SEBN-TN Maintenance Suite** est une plateforme web intégrée dédiée au pilotage complet de la maintenance industrielle :

1. **PMA (Preventive Maintenance Application)** : Gestion du planning de maintenance préventive, suivi des fiches de contrôle, checklists dynamiques et indicateurs de réalisation.
2. **IMA (Intervention Management Application)** : Déclaration, assignation et clôture des interventions curatives, calcul des temps d'arrêt, analyse Pareto des pannes et indicateurs MTTR / MTBF.
3. **EBM (Equipment Budget Management)** : Suivi budgétaire des équipements, validation des commandes (PO), répartition financière par projet et courbes de tendance.
4. **Passation de Shift** : Registre numérique des passations d'équipe, questionnaires personnalisés et historique des relèves.
5. **Gestion Documentaire Intégrée** : Bibliothèque documentaire centralisée avec visionneuse Excel/PDF en direct dans le navigateur et téléchargement sécurisé.
6. **Contrôle d'Accès Basé sur les Rôles (RBAC)** : Niveaux d'accès granulaires pour Propriétaire (`OWNER`), Administrateurs (`ADMIN`), et Techniciens (`TECHNICIAN`).

---

## 🛠️ Stack Technologique

- **Backend** : Python 3.10+ / Flask / Werkzeug
- **Base de Données** : SQLite avec tables normalisées et migrations automatiques
- **Traitement de Données** : Pandas, OpenPyXL, ReportLab
- **Frontend** : HTML5 sémantique, CSS3 (Design System Corporate SEBN-TN), JavaScript Vanilla, Chart.js, SheetJS
- **Sécurité** : Hachage PBKDF2-SHA256, sessions éphémères HttpOnly, protection contre l'inactivité et les injections de chemin (path traversal).

---

## 🚀 Installation & Démarrage

### 1. Prérequis
- Python 3.10 ou supérieur
- Gestionnaire de paquets `pip`

### 2. Cloner le Répertoire
```bash
git clone https://github.com/your-username/sebn-tn-maintenance.git
cd sebn-tn-maintenance
```

### 3. Créer un Environnement Virtuel
```bash
# Windows
python -m venv .venv
.venv\Scripts\activate

# Linux / MacOS
python3 -m venv .venv
source .venv/bin/activate
```

### 4. Installer les Dépendances
```bash
pip install -r requirements.txt
```

### 5. Initialiser la Configuration et la Base de Données
```bash
cp .env.example .env
python scripts/init_db.py
```

### 6. Lancer le Serveur d'Application
```bash
python run.py
```

L'application sera accessible sur :
- **Local** : `http://localhost:5000`
- **Réseau local** : `http://<IP_MACHINE>:5000`

---

## 🧪 Exécution des Tests

Le projet intègre une suite complète de tests unitaires et d'intégration :

```bash
# Lancer tous les tests avec pytest
pytest tests/ -v
```

Les tests vérifient :
- L'authentification et la gestion des sessions
- Le contrôle d'accès basé sur les rôles (RBAC)
- Les interventions curatives (IMA)
- La planification préventive (PMA)
- La gestion documentaire et la visionneuse Excel
- L'intégrité et la sécurité des données

---

## 👥 Rôles & Accès

| Rôle | Accès & Permissions |
|---|---|
| **Propriétaire (`OWNER`)** | Accès total : EBM, Passation, Paramètres système, Gestion des utilisateurs, Réinitialisation sécurisée des accès. |
| **Administrateur (`ADMIN`)** | Gestion des interventions, validation des checklists, attribution des formations, gestion documentaire. |
| **Technicien (`TECHNICIAN`)** | Saisie des checklists PMA, déclaration/clôture des pannes, remplissage des passations de shift, consultation des documents. |

---

## 📄 Licence & Droits d'Auteur

© SEBN-TN — Système Intégré de Maintenance Industrielle.  
**Développé par Mohamed Aissaoui**
