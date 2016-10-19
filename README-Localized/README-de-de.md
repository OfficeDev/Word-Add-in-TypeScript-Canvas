# <a name="image-callouts-word-add-in-sample:-load,-edit,-and-insert-images"></a>Word-Add-In-Beispiel für Bildbeschriftungen: Laden, Bearbeiten und Einfügen von Bildern

**Inhaltsverzeichnis**

* [Zusammenfassung](#summary)
* [Erforderliche Tools](#required-tools)
* [Installieren von Zertifikaten](#how-to-install-certificates)
* [Einrichten und Ausführen der App](#how-to-set-up-and-run-the-app)
* [Ausführen des Add-Ins in Word 2016 unter Windows](#how-to-run-the-add-in-in-Word-2016-for-Windows)
* [Häufig gestellte Fragen](#faq)
* [Fragen und Kommentare](#questions-and-comments)
* [Weitere Informationen](#learn-more)


## <a name="summary"></a>Zusammenfassung

In diesem Word-Add-in-Beispiel wird gezeigt, wie folgende Aktionen ausführen:

1. Erstellen eines Word-Add-Ins mit TypeScript
2. Laden von Bildern aus dem Dokument in das Add-In
3. Bearbeiten von Bildern im Add-In mithilfe der HTML Canvas-API und Einfügen von Bildern im Word-Dokument
4. Implementieren von Add-In-Befehlen zum Starten eines Add-Ins im Menüband und zum Ausführen eines Skripts im Menüband und über ein Kontextmenü
5. Verwenden der Office-UI-Fabric zum Erstellen einer systemeigenen Word-Umgebung für das Add-In

![](/readme-images/Word-Add-in-TypeScript-Canvas.gif)

Definition eines **Add-In-Befehls**: Eine Erweiterung der Word-Benutzeroberfläche, mit dem Sie das Add-In in einem Aufgabenbereich starten oder ein Skript im Menüband oder über ein Kontextmenü ausführen können.

Wenn Sie dies in Aktion sehen möchten, fahren Sie mit [Einrichten von Word 2016 unter Windows](#word-2016-for-windows-set-up) fort, und verwenden Sie dieses [Manifest](https://github.com/OfficeDev/Word-Add-in-TypeScript-Canvas/blob/deploy2Azure/manifest-word-add-in-canvas.xml).

## <a name="required-tools"></a>Erforderliche Tools

Für die Verwendung des Word-Add-In-Beispiels für Bildbeschriftungen gelten folgende Anforderungen.

* Word 2016 16.0.6326.0000 oder höher oder beliebiger Client, der die Javascript-API für Word unterstützt. In diesem Beispiel wird eine Prüfung der Anforderungen durchgeführt, um zu sehen, ob dieses in einem unterstützten Host für die JavaScript-APIs ausgeführt wird.
* npm (https://www.npmjs.com/) zum Installieren der Abhängigkeiten. Es ist im Lieferumfang von [NodeJS](https://nodejs.org/en/) enthalten.
* (Windows) [Git Bash](http://www.git-scm.com/downloads).
* Klonen Sie dieses Repository auf ihrem lokalen Computer.

> Hinweis: Word für Mac 2016 unterstützt derzeit keine Add-In-Befehle. Dieses Beispiel kann auf dem Mac ohne die Add-In-Befehle ausgeführt werden.

## <a name="how-to-install-certificates"></a>Installieren von Zertifikaten

Sie benötigen ein Zertifikat, um dieses Beispiel auszuführen. Da Add-In-Befehle HTTPS erfordern und ohne Oberfläche sind, können Sie keine ungültigen Zertifikate akzeptieren. Führen Sie [./gen-cert.sh](#gen-cert.sh) aus, um das Zertifikat zu erstellen, und installieren Sie dann „ca.crt“ im Speicher mit vertrauenswürdigen Stammzertifizierungsstellen (Windows).

## <a name="how-to-set-up-and-run-the-app"></a>Einrichten und Ausführen der App

1. Installieren Sie den TypeScript-Definitions-Manager, indem Sie ```npm install typings -g``` in der Befehlszeile eingeben.
2. Installieren Sie die in der Datei „typings.json“ identifizierten TypeScript-Definitionen, indem Sie ```typings install``` im Stammverzeichnis des Projekts in der Befehlszeile ausführen.
3. Installieren Sie die in der Datei „package.json“ identifizierten Projektabhängigkeiten, indem Sie ```npm install``` im Stammverzeichnis des Projekts ausführen.
4. Installieren Sie gulp ```npm install -g gulp```.
5. Kopieren Sie die Fabric- und JQuery-Dateien, indem Sie ```gulp copy:libs``` ausführen. (Windows) Wenn dabei ein Problem auftritt, stellen Sie sicher, dass *%APPDATA%\npm* in der path-Variablen enthalten ist.
6. Führen Sie die gulp-Standardaufgabe aus, indem Sie ```gulp``` im Stammverzeichnis des Projekts ausführen. Wenn die TypeScript-Definitionen nicht aktualisiert sind, tritt dabei ein Fehler auf.

Sie haben nun dieses Beispiel-Add-In bereitgestellt. Jetzt müssen Sie Word mitteilen, wo es das Add-In finden kann.

### <a name="word-2016-for-windows-set-up"></a>Einrichtung von Word 2016 unter Windows

1. (Nur Windows) Extrahieren und führen Sie diesen [Registrierungsschlüssel](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/tree/master/Tools/AddInCommandsUndark) aus, um die Add-In-Befehle zu aktivieren. Dies ist erforderlich, solange Add-In-Befehle eine **Vorschaufunktion** sind.
2. Erstellen Sie eine Netzwerkfreigabe oder [geben Sie einen Ordner im Netzwerk frei](https://technet.microsoft.com/en-us/library/cc770880.aspx), und platzieren Sie die [manifest-word-add-in-canvas.xml](manifest-word-add-in-canvas.xml)-Manifestdatei darin.
3. Starten Sie Word, und öffnen Sie ein Dokument.
4. Klicken Sie auf die Registerkarte **Datei**, und klicken Sie dann auf **Optionen**.
5. Wählen Sie **Sicherheitscenter** aus, und klicken Sie dann auf die Schaltfläche **Einstellungen für das Sicherheitscenter**.
6. Wählen Sie **Vertrauenswürdige Add-In-Kataloge** aus.
7. Geben Sie in das Feld **Katalog-URL** den Netzwerkpfad zur Ordnerfreigabe an, die die Datei „manifest-word-add-in-canvas.xml“ enthält, und wählen Sie dann **Katalog hinzufügen**.
8. Aktivieren Sie das Kontrollkästchen **Im Menü anzeigen**, und klicken Sie dann auf **OK**.
9. Es wird eine Meldung angezeigt, dass Ihre Einstellungen angewendet werden, wenn Sie Office das nächste Mal starten.

## <a name="how-to-run-the-add-in-in-word-2016-for-windows"></a>Ausführen des Add-Ins in Word 2016 unter Windows

1. Öffnen Sie ein Word-Dokument.
2. Klicken Sie auf der Registerkarte **Einfügen** in Word 2016 auf **Meine-Add-Ins**.
3. Klicken Sie auf die Registerkarte **Freigegebener Ordner**.
4. Wählen Sie das **Add-In für Bildbeschriftungen**, und wählen Sie dann **Einfügen**.
5. Wenn Add-In-Befehle von Ihrer Word-Version unterstützt werden, werden Sie in der Benutzeroberfläche darüber informiert, dass das Add-In geladen wurde. Sie können die Registerkarte **Add-In für Beschriftungen** verwenden, um das Add-In in der Benutzeroberfläche zu laden und ein Bild in das Dokument einzufügen. Sie können auch mit der rechten Maustaste im Kontextmenü klicken, um ein Bild in das Dokument einzufügen.
6. Wenn Add-In-Befehle von Ihrer Version von Word nicht unterstützt werden, wird das Add-In in einem Aufgabenbereich geladen. Sie müssen ein Bild in das Word-Dokument einfügen, um diese Funktion des Add-Ins zu verwenden.
7. Wählen Sie ein Bild im Word-Dokument, und laden Sie es im Aufgabenbereich durch Wählen von *Bild aus Dokument laden*. Jetzt können Sie Beschriftungen im Bild einfügen. Wählen Sie *Bild im Dokument einfügen*, um das aktualisierte Bild im Word-Dokument einzufügen. Das Add-In generiert Platzhalterbeschreibungen für alle Beschriftungen.

## <a name="faq"></a>Häufig gestellte Fragen

* Funktionieren Add-In-Befehle auf dem Mac und iPad? Nein, sie funktionieren seit der Veröffentlichung dieser Readme-Datei nicht auf dem Mac oder iPad.
* Warum wird mein Add-In nicht im Fenster **Meine Add-Ins** angezeigt? Möglicherweise weist das Add-In-Manifest einen Fehler auf. Es wird empfohlen, die Manifestdatei anhand des [Manifestdateischemas](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/tree/master/Tools/XSD) zu prüfen.
* Warum wird die Funktionsdatei nicht für meine Add-In-Befehle aufgerufen? Add-In-Befehle erfordern HTTPS. Da Add-In-Befehle TLS erfordern une es keine Benutzeroberfläche gibt, können Sie nicht sehen, ob ein Problem mit dem Zertifikat besteht. Wenn Sie ein ungültiges Zertifikat im Aufgabenbereich akzeptieren müssen, funktioniert der Add-In-Befehl nicht.
* Warum reagieren die npm-Installationsbefehle nicht? Dies trifft wahrscheinlich nicht zu. Dies nimmt unter Windows lediglich einige Zeit in Anspruch.

## <a name="questions-and-comments"></a>Fragen und Kommentare

Wir schätzen Ihr Feedback hinsichtlich des Word-Add-In-Beispiels für Bildbeschriftungen. Sie können uns Ihre Fragen und Vorschläge über den Abschnitt [Probleme](https://github.com/OfficeDev/Word-Add-in-TypeScript-Canvas/issues) dieses Repositorys senden.

Allgemeine Fragen zur Add-In-Entwicklung sollten in [Stack Overflow](http://stackoverflow.com/questions/tagged/Office365+API) gestellt werden. Stellen Sie sicher, dass Ihre Fragen oder Kommentare mit [office-js], [word-addins] und [API] markiert sind. Diese Markierunge werden von uns überprüft.

## <a name="learn-more"></a>Weitere Informationen

Hier noch einige weitere Ressourcen zum Erstellen von Add-Ins auf Basis von Word-JavaScript-APIs:

* [Word-Add-Ins – Dokumentation und Beispiele](https://dev.office.com/word)
* [Beispiel für SillyStories](https://github.com/OfficeDev/Word-Add-in-SillyStories) – Informationen zum Laden von DOCX-Dateien aus einem Dienst und Einfügen der Dateien in ein geöffnetes Word-Dokument
* [Serverauthentifizierungsbeispiel für Office-Add-In für Node.js](https://github.com/OfficeDev/Office-Add-in-Nodejs-ServerAuth) Informationen zum Verwenden von Azure- und Google OAuth-Anbieter zum Authentifizieren von Add-In-Benutzern

## <a name="copyright"></a>Copyright
Copyright (c) 2016 Microsoft. Alle Rechte vorbehalten.
