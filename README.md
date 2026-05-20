# Suivi Activités Labo

Application Windows WPF de suivi des projets et essais laboratoire.

Version actuelle : 1.0

## Fonctionnalités

- Tableau général des projets avec recherche, tri, filtres et édition directe.
- Ajout et modification de projets avec références produit multiples séparées par des virgules.
- Suivi des essais par projet, avec statuts, résultat OK/NOK, répétitions et référence produit utilisée.
- Vue des projets en cours avec synthèse d’avancement, essais faits et essais à suivre.
- Agenda de travail hebdomadaire avec planification des essais par glisser-déposer, pauses, jours travaillés et durées configurables.
- Gestion des propriétés de référence : clients, demandeurs, familles produit, types d’activité, essais, catégories, durées et statuts disponibles.
- Export Excel `.xlsx` des projets.
- Archivage/désarchivage des projets terminés.

## Données

Les données applicatives sont stockées dans `data.json`, copié à côté de l’exécutable lors du build ou de la publication.

Le fichier `reference-properties.json`, lorsqu’il existe à côté de l’exécutable, contient les listes de référence configurables utilisées par l’application.

## Prérequis développement

- Windows
- .NET SDK 8 ou plus récent

## Build

```powershell
dotnet build suiviActivites/suiviActivites.csproj
```

## Publication self-contained

La publication self-contained embarque le runtime .NET pour Windows x64.

```powershell
dotnet publish suiviActivites/suiviActivites.csproj -c Release -r win-x64 --self-contained true -p:PublishSingleFile=false -o publish/suiviActivites-1.0-self-contained
```

Le point d’entrée est :

```text
publish/suiviActivites-1.0-self-contained/suiviActivites.exe
```

## Structure principale

- `suiviActivites/Models` : modèles métier des projets, essais et agenda.
- `suiviActivites/ViewModels/MainViewModel.cs` : logique principale, sauvegarde, filtres, statistiques, agenda et propriétés.
- `suiviActivites/Views` : vues WPF de l’application.
- `suiviActivites/Services/XlsxExportService.cs` : génération de l’export Excel.
- `publish/` : dossiers et archives de publication.

