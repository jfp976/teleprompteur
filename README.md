# Téléprompteur

Téléprompteur pour Windows : le texte défile en gros caractères Arial, avec mode karaoké synchronisé sur une voix de synthèse.
Langues de l'interface et de la voix : français, anglais, espagnol, arabe (lecture et karaoké de droite à gauche).

## Télécharger

La dernière version est dans **[Releases](../../releases)** : `Teleprompteur.exe` fonctionne seul, sans installer Python.
Au premier lancement, Windows SmartScreen peut avertir (exécutable non signé) : « Informations complémentaires » → « Exécuter quand même ».

## Raccourcis

| Touche | Action |
|---|---|
| Espace | lecture / pause |
| Molette | vitesse de défilement (vitesse du karaoké quand il est actif) |
| ↑ / ↓ | paragraphe précédent / suivant |
| Début / Fin | début / fin du texte |
| Page préc. / Page suiv. | écran précédent / suivant |
| + / − | taille des caractères |
| * | mode karaoké (le texte lu passe en jaune) |
| ← / → | karaoké moins vite / plus vite |
| Clic | modifier le texte (Échap pour revenir au défilement) |
| F1 | aide |
| F2 | voix de synthèse on / off |
| F3 | ouvrir un fichier .txt, .pdf, .doc, .docx |
| F4 | langue : Français → English → Español → العربية |
| F11 | plein écran |
| Échap | quitter (l'enregistrement des modifications est proposé) |

## Voix

Voix neuronales Microsoft (Denise, Sonia, Elvira, Zariyah) si Internet est disponible, sinon les voix installées dans Windows (Virginie, Hortense…).

## Lancer depuis les sources

Python 3.10+ requis. Double-cliquer sur `Lancer_teleprompteur.bat` (installe les modules de `requirements.txt`).
Pour reconstruire l'exécutable : `Construire_exe.bat`.

## Publier une nouvelle version

Changer le numéro dans le fichier `VERSION` (par exemple `v1.1`) et pousser sur `main` :
GitHub construit `Teleprompteur.exe` sous Windows et crée la release correspondante.
