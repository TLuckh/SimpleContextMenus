﻿- NamingConventionParser dürfte momentan noch für Ordner failen?
- Hinzufügen von Bildern an die Kontextmenü-Einträge
- - (Ein einfaches .bmp mit gleichem Namen wie Ordner/Datei sollte es tun)
- Andere Skriptsprachen und .exe's sind theoretisch bereits möglich, aber solten dokumentiert werden.
- Restriktionen auf Ordner
- Erhöhen der Effizienz:
- - Statt das Kontextmenü jedes Mal neu zu bauen, können wir es anfangs einmal bauen und dann bei jedem Aufruf nur noch die Sichtbarkeit ändern. (MenuStripItem.Visible)
- - Jedes Dropdown-Kontextmenü könnte Lazy-loaded werden.
- 